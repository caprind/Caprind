VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Pedido 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Pedido"
   ClientHeight    =   10035
   ClientLeft      =   1305
   ClientTop       =   1650
   ClientWidth     =   15360
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
   Icon            =   "frmCompras_Pedido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   690
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   209
      Top             =   9735
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
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
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmCompras_Pedido.frx":014A
      Left            =   12630
      List            =   "frmCompras_Pedido.frx":014C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      ToolTipText     =   "Empresa."
      Top             =   1810
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   75
      TabIndex        =   210
      Top             =   8880
      Visible         =   0   'False
      Width           =   15195
      Begin VB.TextBox txt_baseICMS_ST 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Base de calculo ICMS substituto."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txt_VlrICMS 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do ICMS."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txt_BaseICMS 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Base de cálculo do ICMS."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txtTotalSeguro 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8610
         MaxLength       =   50
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do seguro."
         Top             =   390
         Width           =   1135
      End
      Begin VB.TextBox txtTotalAcessorias 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9765
         MaxLength       =   50
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total das despesas acessórias."
         Top             =   390
         Width           =   1080
      End
      Begin VB.TextBox txtTotalFrete 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7458
         MaxLength       =   50
         TabIndex        =   67
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do frete."
         Top             =   390
         Width           =   1135
      End
      Begin VB.TextBox txt_TotalIPI 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11247
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do IPI."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txt_VlrTotalProd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5025
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do(s) produto(s)."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txtTotalServicos 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do(s) serviço(s)."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txtTotalPedido 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   13830
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do pedido."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txtTotalDesconto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   12465
         MaxLength       =   50
         TabIndex        =   71
         Text            =   "0,00"
         ToolTipText     =   "Valor total do desconto."
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox txt_ICMS_ST 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do ICMS substituto."
         Top             =   390
         Width           =   1185
      End
      Begin DrawSuite2022.USButton cmdSalvar_Frete 
         Height          =   315
         Left            =   10860
         TabIndex        =   408
         ToolTipText     =   "Salvar o valor total do frete, seguro e despesas acessórias"
         Top             =   390
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Pedido.frx":014E
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin DrawSuite2022.USButton cmdSalvar_desconto 
         Height          =   315
         Left            =   13440
         TabIndex        =   409
         Top             =   390
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Pedido.frx":8B53
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total pedido"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   90
         Left            =   13890
         TabIndex        =   350
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total desc."
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   89
         Left            =   12495
         TabIndex        =   349
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total IPI"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   88
         Left            =   11475
         TabIndex        =   348
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total desp."
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   87
         Left            =   9840
         TabIndex        =   347
         Top             =   180
         Width           =   810
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total seguro"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   86
         Left            =   8640
         TabIndex        =   346
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total frete"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   85
         Left            =   7575
         TabIndex        =   345
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total serviço"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   84
         Left            =   6285
         TabIndex        =   344
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total prod."
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   83
         Left            =   5160
         TabIndex        =   343
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS ST"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   82
         Left            =   3825
         TabIndex        =   342
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BC ICMS ST"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   81
         Left            =   2730
         TabIndex        =   341
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   80
         Left            =   1530
         TabIndex        =   340
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BC do ICMS"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   211
         Top             =   180
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   301
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Carteira de compras"
      TabPicture(0)   =   "frmCompras_Pedido.frx":11558
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(2)=   "PBLista(0)"
      Tab(0).Control(3)=   "SSTab4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Pedido de compra"
      TabPicture(1)   =   "frmCompras_Pedido.frx":11574
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(0)"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "listapedido"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "Txtidpedido"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1(16)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Dados comerciais"
      TabPicture(2)   =   "frmCompras_Pedido.frx":11590
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(1)=   "Frame1(1)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Lista de produtos"
      TabPicture(3)   =   "frmCompras_Pedido.frx":115AC
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "USToolBar4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "SSTab2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lista de serviços"
      TabPicture(4)   =   "frmCompras_Pedido.frx":115C8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTab3"
      Tab(4).Control(1)=   "USToolBar5"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Escopo de fornecimento"
      TabPicture(5)   =   "frmCompras_Pedido.frx":115E4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1(11)"
      Tab(5).Control(1)=   "USToolBar6"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do fornecedor"
         Enabled         =   0   'False
         Height          =   2055
         Index           =   16
         Left            =   -74925
         TabIndex        =   291
         Top             =   2265
         Width           =   15225
         Begin VB.TextBox txtCNPJ 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   9150
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   411
            ToolTipText     =   "Número."
            Top             =   375
            Width           =   1695
         End
         Begin VB.TextBox Txt_descricao_referencia 
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
            Height          =   315
            Left            =   10740
            MaxLength       =   30
            TabIndex        =   60
            ToolTipText     =   "Descrição da referência."
            Top             =   1605
            Width           =   4290
         End
         Begin VB.TextBox Txt_n_referencia 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   9280
            MaxLength       =   30
            TabIndex        =   59
            ToolTipText     =   "Número da referência."
            Top             =   1605
            Width           =   1440
         End
         Begin VB.TextBox txtEmail 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   56
            ToolTipText     =   "E-mail."
            Top             =   1605
            Width           =   5670
         End
         Begin VB.TextBox txtcontato 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   10845
            MaxLength       =   255
            TabIndex        =   48
            ToolTipText     =   "Nome de contato."
            Top             =   375
            Width           =   3840
         End
         Begin VB.TextBox txtfornecedor 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   735
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Nome do fornecedor."
            Top             =   375
            Width           =   7665
         End
         Begin VB.TextBox txtIDfornecedor 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   180
            TabIndex        =   45
            ToolTipText     =   "Código do fornecedor."
            Top             =   375
            Width           =   540
         End
         Begin VB.TextBox txtendereco 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1215
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   50
            ToolTipText     =   "Endereço."
            Top             =   1005
            Width           =   4500
         End
         Begin VB.TextBox txtCidade 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   11370
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   54
            ToolTipText     =   "Cidade."
            Top             =   1005
            Width           =   3135
         End
         Begin VB.TextBox txtbairro 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   7785
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   53
            ToolTipText     =   "Bairro."
            Top             =   1005
            Width           =   3570
         End
         Begin VB.TextBox txttelefone 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   5865
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   57
            ToolTipText     =   "Número do telefone."
            Top             =   1605
            Width           =   1695
         End
         Begin VB.TextBox txtuf 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   14520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   55
            ToolTipText     =   "UF."
            Top             =   1005
            Width           =   510
         End
         Begin VB.TextBox txtfax 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   7575
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   58
            ToolTipText     =   "Número do fax."
            Top             =   1605
            Width           =   1695
         End
         Begin VB.TextBox txtNumero 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   5730
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   51
            ToolTipText     =   "Número."
            Top             =   1005
            Width           =   1005
         End
         Begin VB.TextBox txtTipo_bairro 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   6750
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   52
            ToolTipText     =   "Bairro."
            Top             =   1005
            Width           =   1020
         End
         Begin VB.TextBox txtTipo_endereco 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   49
            ToolTipText     =   "Bairro."
            Top             =   1005
            Width           =   1020
         End
         Begin VB.TextBox txtCategoria 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   47
            ToolTipText     =   "Categoria do fornecedor."
            Top             =   375
            Width           =   390
         End
         Begin DrawSuite2022.USButton cmdAdicionarfornecedor 
            Height          =   315
            Left            =   8820
            TabIndex        =   399
            Top             =   360
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":11600
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
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin DrawSuite2022.USButton cmdcontatos 
            Height          =   315
            Left            =   14700
            TabIndex        =   400
            Top             =   360
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":2F705
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
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "CPF | CNPJ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   130
            Left            =   9630
            TabIndex        =   412
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição da referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   79
            Left            =   12038
            TabIndex        =   339
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "N. da referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   78
            Left            =   9415
            TabIndex        =   338
            Top             =   1410
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   77
            Left            =   6397
            TabIndex        =   337
            Top             =   1410
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   76
            Left            =   8287
            TabIndex        =   336
            Top             =   1410
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   75
            Left            =   9360
            TabIndex        =   335
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   74
            Left            =   7110
            TabIndex        =   334
            Top             =   810
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   73
            Left            =   5955
            TabIndex        =   333
            Top             =   810
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   3128
            TabIndex        =   332
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   71
            Left            =   14678
            TabIndex        =   331
            Top             =   810
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Contato"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   70
            Left            =   12473
            TabIndex        =   330
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cat."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   8475
            TabIndex        =   329
            Top             =   180
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   12690
            TabIndex        =   321
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome razão social"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   3922
            TabIndex        =   295
            Top             =   180
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   570
            TabIndex        =   293
            Top             =   810
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   2805
            TabIndex        =   292
            Top             =   1410
            Width           =   420
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Empresa"
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
         Height          =   675
         Left            =   -74865
         TabIndex        =   302
         Top             =   1620
         Width           =   4695
         Begin VB.ComboBox Cmb_empresa_carteira 
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
            Height          =   330
            ItemData        =   "frmCompras_Pedido.frx":4D80A
            Left            =   210
            List            =   "frmCompras_Pedido.frx":4D811
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   240
            Width           =   4305
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
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
            Left            =   1980
            TabIndex        =   303
            Top             =   210
            Width           =   765
         End
      End
      Begin VB.TextBox Txtidpedido 
         Height          =   375
         Left            =   -73140
         TabIndex        =   296
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   6270
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   283
         Top             =   8280
         Width           =   15225
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
            Index           =   3
            Left            =   9540
            TabIndex        =   76
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
            Index           =   3
            Left            =   2880
            TabIndex        =   74
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.ComboBox Cmb_opcao_lista 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmCompras_Pedido.frx":4D822
            Left            =   6840
            List            =   "frmCompras_Pedido.frx":4D82C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   180
            Width           =   1965
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   3
            Left            =   11760
            TabIndex        =   80
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":4D846
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagAnt 
            Height          =   315
            Index           =   3
            Left            =   11220
            TabIndex        =   79
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":50FEA
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagIr 
            Height          =   315
            Index           =   3
            Left            =   10110
            TabIndex        =   77
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagPrim 
            Height          =   315
            Index           =   3
            Left            =   10680
            TabIndex        =   78
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":54AF3
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin DrawSuite2022.USButton cmdPagUlt 
            Height          =   315
            Index           =   3
            Left            =   12300
            TabIndex        =   81
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":58BE2
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
            BorderColor     =   14404026
            BorderColorDown =   11632444
            BorderColorOver =   11632444
            GradientColor2  =   16777215
            GradientColor3  =   16777215
            GradientColorOver1=   16643560
            GradientColorOver2=   16576988
            GradientColorOver3=   16441780
            GradientColorOver4=   16178091
            GradientColorDown2=   16246986
            GradientColorDown3=   15189380
            GradientColorDown4=   14596208
            PicSizeH        =   19
            PicSizeW        =   19
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   128
            Left            =   3510
            TabIndex        =   392
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   13050
            TabIndex        =   287
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   286
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   2190
            TabIndex        =   285
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   5520
            TabIndex        =   284
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   8685
         Index           =   1
         Left            =   -74925
         TabIndex        =   271
         Top             =   1305
         Width           =   15225
         Begin VB.CommandButton cmdDadosComerciaisFornecedor 
            Caption         =   "Buscar dados comerciais do fornecedor"
            Height          =   435
            Left            =   2250
            TabIndex        =   396
            ToolTipText     =   "Buscar dados comerciais do cadasto do fornecedor"
            Top             =   7140
            Width           =   12765
         End
         Begin VB.TextBox Txt_ID_entrega 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2250
            TabIndex        =   282
            Text            =   "0"
            Top             =   4696
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtembalagem 
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
            Height          =   1090
            Left            =   2250
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   84
            TabStop         =   0   'False
            ToolTipText     =   "Embalagem."
            Top             =   2438
            Width           =   12405
         End
         Begin VB.TextBox txtprazo 
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
            Height          =   1090
            Left            =   2250
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "Prazo de entrega."
            Top             =   1309
            Width           =   12405
         End
         Begin VB.TextBox txtobservacoes 
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
            Height          =   1305
            Left            =   2265
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            ToolTipText     =   "Observações."
            Top             =   5775
            Width           =   12735
         End
         Begin VB.TextBox cmbpagamento 
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
            Height          =   1090
            Left            =   2250
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "Condições de pagamento."
            Top             =   180
            Width           =   12405
         End
         Begin VB.ComboBox cmbtransporte 
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
            ItemData        =   "frmCompras_Pedido.frx":5C46E
            Left            =   4980
            List            =   "frmCompras_Pedido.frx":5C470
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   88
            ToolTipText     =   "Transportadora."
            Top             =   5050
            Width           =   5115
         End
         Begin VB.TextBox txtlocal 
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
            Height          =   315
            Left            =   2250
            MaxLength       =   255
            TabIndex        =   86
            ToolTipText     =   "Local de entrega."
            Top             =   4696
            Width           =   12405
         End
         Begin VB.TextBox txtBanco 
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
            Height          =   315
            Left            =   2250
            MaxLength       =   255
            TabIndex        =   91
            ToolTipText     =   "Banco."
            Top             =   5419
            Width           =   5025
         End
         Begin VB.TextBox txtConta 
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
            Height          =   315
            Left            =   11805
            MaxLength       =   255
            TabIndex        =   93
            ToolTipText     =   "Conta corrente."
            Top             =   5415
            Width           =   2865
         End
         Begin VB.TextBox txtAgencia 
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
            Height          =   315
            Left            =   8130
            MaxLength       =   255
            TabIndex        =   92
            ToolTipText     =   "Agência."
            Top             =   5415
            Width           =   2445
         End
         Begin VB.TextBox txttransporte 
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
            Height          =   1090
            Left            =   2250
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   85
            TabStop         =   0   'False
            ToolTipText     =   "Transporte."
            Top             =   3567
            Width           =   12405
         End
         Begin VB.ComboBox cmbMoeda 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "frmCompras_Pedido.frx":5C472
            Left            =   11205
            List            =   "frmCompras_Pedido.frx":5C474
            Style           =   2  'Dropdown List
            TabIndex        =   89
            ToolTipText     =   "Moeda."
            Top             =   5050
            Width           =   1695
         End
         Begin VB.TextBox Txt_valor_moeda 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   13975
            MaxLength       =   60
            TabIndex        =   90
            ToolTipText     =   "Valor da moeda."
            Top             =   5050
            Width           =   1010
         End
         Begin VB.ComboBox Cmb_tipo_transp 
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
            Height          =   330
            ItemData        =   "frmCompras_Pedido.frx":5C476
            Left            =   2250
            List            =   "frmCompras_Pedido.frx":5C486
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   87
            ToolTipText     =   "Tipo da transportadora."
            Top             =   5050
            Width           =   1335
         End
         Begin VB.CheckBox chkObs_Financeiro 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Salvar observação no financeiro"
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
            Left            =   12000
            TabIndex        =   95
            Top             =   7560
            Width           =   3015
         End
         Begin DrawSuite2022.USButton cmdCondicoes 
            Height          =   1095
            Left            =   14670
            TabIndex        =   402
            Top             =   180
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   1931
            DibPicture      =   "frmCompras_Pedido.frx":5C4AA
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton cmdEntrega 
            Height          =   1095
            Left            =   14670
            TabIndex        =   403
            Top             =   1320
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   1931
            DibPicture      =   "frmCompras_Pedido.frx":7A5AF
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton cmdEmbalagem 
            Height          =   1095
            Left            =   14670
            TabIndex        =   404
            Top             =   2460
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   1931
            DibPicture      =   "frmCompras_Pedido.frx":986B4
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton cmdTransporte_padrao 
            Height          =   1095
            Left            =   14670
            TabIndex        =   405
            Top             =   3570
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   1931
            DibPicture      =   "frmCompras_Pedido.frx":B67B9
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton cmdLocTransp 
            Height          =   315
            Left            =   10110
            TabIndex        =   406
            Top             =   5040
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":D48BE
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
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin DrawSuite2022.USButton cmdDados_pagto 
            Height          =   315
            Left            =   14670
            TabIndex        =   407
            Top             =   5400
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":F29C3
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
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin DrawSuite2022.USButton cmdlocalentrega 
            Height          =   315
            Left            =   14670
            TabIndex        =   398
            Top             =   4680
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":110AC8
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Ct. corrente :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   94
            Left            =   10740
            TabIndex        =   354
            Top             =   5415
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Agência :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   91
            Left            =   7380
            TabIndex        =   353
            Top             =   5415
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Moeda :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   93
            Left            =   10545
            TabIndex        =   352
            Top             =   5055
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Transportadora :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   92
            Left            =   3690
            TabIndex        =   351
            Top             =   5055
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   1260
            TabIndex        =   280
            Top             =   3570
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Condições de pagamento :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   279
            Top             =   180
            Width           =   1920
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo de entrega :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   810
            TabIndex        =   278
            Top             =   1309
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Embalagem :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   1245
            TabIndex        =   277
            Top             =   2438
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de entrega :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   855
            TabIndex        =   276
            Top             =   4696
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo da transportadora :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   390
            TabIndex        =   275
            Top             =   5055
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   1620
            TabIndex        =   274
            Top             =   5415
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Observação :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   1185
            TabIndex        =   273
            Top             =   5775
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. moeda :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   13050
            TabIndex        =   272
            Top             =   5055
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H8000000E&
         Height          =   8415
         Index           =   11
         Left            =   -74925
         TabIndex        =   212
         Top             =   1305
         Width           =   15225
         Begin VB.TextBox txtEscopo 
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
            Height          =   7455
            Left            =   180
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   294
            ToolTipText     =   "Escopo de fornecimento."
            Top             =   240
            Width           =   14835
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar6 
         Height          =   975
         Left            =   -74925
         TabIndex        =   213
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   10
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Novo (Insert)"
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
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Filtrar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Filtrar (F2)"
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
         ButtonWidth2    =   36
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Salvar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Salvar (F3)"
         ButtonKey3      =   "3"
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
         ButtonLeft3     =   78
         ButtonTop3      =   2
         ButtonWidth3    =   44
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
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
         ButtonLeft4     =   124
         ButtonTop4      =   2
         ButtonWidth4    =   60
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   186
         ButtonTop5      =   2
         ButtonWidth5    =   55
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   243
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonAlignment7=   2
         ButtonType7     =   1
         ButtonStyle7    =   -1
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState7    =   -1
         ButtonLeft7     =   300
         ButtonTop7      =   4
         ButtonWidth7    =   2
         ButtonHeight7   =   54
         ButtonCaption8  =   "Ajuda"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Ajuda (F1)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   304
         ButtonTop8      =   2
         ButtonWidth8    =   41
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Sair"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Sair (Esc)"
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   347
         ButtonTop9      =   2
         ButtonWidth9    =   30
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   5
         ButtonLeft10    =   379
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         Begin DrawSuite2022.USImageList USImageList6 
            Left            =   13950
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Pedido.frx":114118
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   8685
         Left            =   -74925
         TabIndex        =   214
         Top             =   1350
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15319
         _Version        =   393216
         TabHeight       =   520
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Dados do serviço"
         TabPicture(0)   =   "frmCompras_Pedido.frx":119387
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ListaServ"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtIDLista_serv"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtIDcarteira_serv"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtcodproduto_serv"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Centro de custo"
         TabPicture(1)   =   "frmCompras_Pedido.frx":1193A3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(8)"
         Tab(1).Control(1)=   "txtIDCentro_serv"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame1(9)"
         Tab(1).Control(3)=   "Lista_custoServ"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Empenhos"
         TabPicture(2)   =   "frmCompras_Pedido.frx":1193BF
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(10)"
         Tab(2).Control(1)=   "Lista_empenhos_serv"
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Index           =   10
            Left            =   -74940
            TabIndex        =   299
            Top             =   6660
            Width           =   15135
            Begin VB.TextBox Txt_qtde_total_disp 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   13380
               Locked          =   -1  'True
               TabIndex        =   208
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade disponível."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_emp 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   11490
               Locked          =   -1  'True
               TabIndex        =   207
               TabStop         =   0   'False
               ToolTipText     =   "Quatidade total empenhada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_comprada 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   9660
               Locked          =   -1  'True
               TabIndex        =   206
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade total comprada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. disponível"
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
               Index           =   125
               Left            =   13492
               TabIndex        =   388
               Top             =   210
               Width           =   1350
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. empenhada"
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
               Index           =   124
               Left            =   11527
               TabIndex        =   387
               Top             =   210
               Width           =   1500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. comprada"
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
               Index           =   56
               Left            =   9772
               TabIndex        =   386
               Top             =   210
               Width           =   1350
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "-                                       ="
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
               Index           =   55
               Left            =   11310
               TabIndex        =   300
               Top             =   480
               Width           =   1965
            End
         End
         Begin VB.TextBox txtcodproduto_serv 
            Height          =   315
            Left            =   1170
            TabIndex        =   238
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4260
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtIDcarteira_serv 
            Enabled         =   0   'False
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
            Left            =   1920
            MouseIcon       =   "frmCompras_Pedido.frx":1193DB
            MousePointer    =   99  'Custom
            TabIndex        =   237
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4260
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtIDLista_serv 
            Height          =   315
            Left            =   450
            TabIndex        =   236
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4260
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   600
            Index           =   8
            Left            =   -74940
            TabIndex        =   234
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtValorCentro_Serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   11295
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   198
               TabStop         =   0   'False
               ToolTipText     =   "Valor."
               Top             =   180
               Width           =   1155
            End
            Begin VB.TextBox txtPercentualCentro_Serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   13785
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   200
               TabStop         =   0   'False
               ToolTipText     =   "Percentual."
               Top             =   180
               Width           =   1155
            End
            Begin VB.CheckBox chkValor_serv 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor :"
               Height          =   255
               Left            =   10530
               TabIndex        =   197
               Top             =   180
               Width           =   765
            End
            Begin VB.CheckBox chkPercentual_serv 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Percentual :"
               Height          =   255
               Left            =   12600
               TabIndex        =   199
               Top             =   180
               Width           =   1185
            End
            Begin VB.ComboBox Cmb_centro_servico 
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
               Left            =   1500
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   196
               ToolTipText     =   "Centro de custo."
               Top             =   180
               Width           =   8910
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Centro de custo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   46
               Left            =   180
               TabIndex        =   235
               Top             =   180
               Width           =   1260
            End
         End
         Begin VB.TextBox txtIDCentro_serv 
            Enabled         =   0   'False
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
            Left            =   -74370
            MouseIcon       =   "frmCompras_Pedido.frx":1196E5
            MousePointer    =   99  'Custom
            TabIndex        =   233
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1650
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   855
            Index           =   9
            Left            =   -74940
            TabIndex        =   232
            Top             =   6660
            Width           =   15135
            Begin VB.TextBox txtTotalCentroServ 
               Alignment       =   1  'Right Justify
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
               Left            =   11790
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCompras_Pedido.frx":1199EF
               MousePointer    =   99  'Custom
               TabIndex        =   203
               TabStop         =   0   'False
               ToolTipText     =   "Valor total centro de custo."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtSaldoCentroServ 
               Alignment       =   1  'Right Justify
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
               Left            =   13380
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCompras_Pedido.frx":119CF9
               MousePointer    =   99  'Custom
               TabIndex        =   204
               TabStop         =   0   'False
               ToolTipText     =   "Saldo."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtVlrTotal_centroServ 
               Alignment       =   1  'Right Justify
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
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCompras_Pedido.frx":11A003
               MousePointer    =   99  'Custom
               TabIndex        =   202
               TabStop         =   0   'False
               ToolTipText     =   "Valor total do serviços."
               Top             =   420
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo"
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
               Index           =   123
               Left            =   13935
               TabIndex        =   385
               Top             =   210
               Width           =   465
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total centro"
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
               Index           =   122
               Left            =   12060
               TabIndex        =   384
               Top             =   210
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total"
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
               Index           =   121
               Left            =   10545
               TabIndex        =   383
               Top             =   210
               Width           =   885
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   3450
            Index           =   7
            Left            =   60
            TabIndex        =   215
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtReferencia_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   3390
               MaxLength       =   255
               TabIndex        =   163
               ToolTipText     =   "Código de referência."
               Top             =   390
               Visible         =   0   'False
               Width           =   3435
            End
            Begin VB.TextBox txtDescricao_comercialServ 
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
               Height          =   465
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   176
               TabStop         =   0   'False
               ToolTipText     =   "Descrição comercial."
               Top             =   1620
               Width           =   7965
            End
            Begin VB.CheckBox Chk_valor_desc2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor do desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4830
               TabIndex        =   188
               Top             =   2820
               Width           =   1365
            End
            Begin VB.CheckBox Chk_desc2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desc. (%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3383
               TabIndex        =   186
               Top             =   2820
               Width           =   1035
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo serviço"
               ForeColor       =   &H00000000&
               Height          =   525
               Index           =   15
               Left            =   11640
               TabIndex        =   217
               Top             =   180
               Width           =   3285
               Begin VB.CheckBox chkManual_serv 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. manual ?"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   1890
                  TabIndex        =   167
                  Top             =   270
                  Width           =   1335
               End
               Begin VB.CheckBox chkAuto_serv 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. automático ?"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   180
                  TabIndex        =   166
                  Top             =   270
                  Width           =   1605
               End
            End
            Begin VB.TextBox txtStatus_serv 
               Alignment       =   2  'Center
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
               Left            =   6840
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   165
               TabStop         =   0   'False
               ToolTipText     =   "Status."
               Top             =   390
               Width           =   4545
            End
            Begin VB.CommandButton cmdAbrir_codigo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2640
               Picture         =   "frmCompras_Pedido.frx":11A30D
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Localizar serviços."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtDesconto_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   3270
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   187
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   3015
               Width           =   1260
            End
            Begin VB.TextBox txtVlrDesconto_serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   4545
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   189
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   3015
               Width           =   1935
            End
            Begin VB.TextBox txtVlrUnitDesc_serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   6495
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   190
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   3015
               Width           =   1885
            End
            Begin VB.TextBox txtObs_serv 
               Alignment       =   2  'Center
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
               Height          =   465
               Left            =   8180
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   177
               ToolTipText     =   "Observações."
               Top             =   1620
               Width           =   6765
            End
            Begin VB.TextBox txtDetalhe_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   8415
               MaxLength       =   50
               TabIndex        =   179
               ToolTipText     =   "Detalhe."
               Top             =   2385
               Width           =   1975
            End
            Begin VB.TextBox txtDescricao_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   168
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   1005
               Width           =   6915
            End
            Begin VB.ComboBox cmbFamilia_serv 
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
               Left            =   180
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   178
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   2385
               Width           =   8220
            End
            Begin VB.ComboBox cmbUn_serv 
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
               ItemData        =   "frmCompras_Pedido.frx":11A40F
               Left            =   180
               List            =   "frmCompras_Pedido.frx":11A411
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   183
               TabStop         =   0   'False
               ToolTipText     =   "Unidade de estoque."
               Top             =   3015
               Width           =   735
            End
            Begin VB.TextBox txtISSQN 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   9855
               MaxLength       =   50
               TabIndex        =   192
               ToolTipText     =   "Valor de % do ISSQN."
               Top             =   3015
               Width           =   1285
            End
            Begin VB.TextBox txtValorUnit_serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1650
               MaxLength       =   50
               TabIndex        =   185
               ToolTipText     =   "Valor unitário."
               Top             =   3015
               Width           =   1605
            End
            Begin VB.CommandButton cmdFiltrar_codigo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2310
               Picture         =   "frmCompras_Pedido.frx":11A413
               Style           =   1  'Graphical
               TabIndex        =   160
               ToolTipText     =   "Filtrar por código interno."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtCodigo 
               Alignment       =   2  'Center
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
               Left            =   180
               MaxLength       =   50
               TabIndex        =   159
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   2115
            End
            Begin VB.TextBox txtQtde_serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   8400
               MaxLength       =   50
               TabIndex        =   191
               ToolTipText     =   "Quantidade."
               Top             =   3015
               Width           =   1440
            End
            Begin VB.TextBox txtValor_ISSQN 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   11160
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   193
               TabStop         =   0   'False
               ToolTipText     =   "Valor do ISSQN."
               Top             =   3015
               Width           =   1885
            End
            Begin VB.TextBox txtValorTotal_serv 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   13070
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   194
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   3015
               Width           =   1875
            End
            Begin VB.Frame framePrazo_serv 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   13520
               TabIndex        =   216
               Top             =   1000
               Width           =   1095
               Begin MSMask.MaskEdBox txtPrazo_serv 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   175
                  ToolTipText     =   "Prazo de entrega."
                  Top             =   0
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   0
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
            End
            Begin VB.ComboBox cmbReferencia_serv 
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
               ItemData        =   "frmCompras_Pedido.frx":11A82E
               Left            =   3390
               List            =   "frmCompras_Pedido.frx":11A830
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   164
               ToolTipText     =   "Código de referencia."
               Top             =   390
               Width           =   3435
            End
            Begin VB.TextBox txtOrdem_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   10410
               TabIndex        =   180
               ToolTipText     =   "Número da ordem de produção."
               Top             =   2385
               Width           =   1735
            End
            Begin VB.ComboBox Cmb_OS_serv 
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
               ItemData        =   "frmCompras_Pedido.frx":11A832
               Left            =   12165
               List            =   "frmCompras_Pedido.frx":11A834
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   181
               ToolTipText     =   "Número da OS."
               Top             =   2385
               Width           =   1185
            End
            Begin VB.ComboBox Cmb_un_com_serv 
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
               ItemData        =   "frmCompras_Pedido.frx":11A836
               Left            =   910
               List            =   "frmCompras_Pedido.frx":11A838
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   184
               TabStop         =   0   'False
               ToolTipText     =   "Unidade comercial."
               Top             =   3015
               Width           =   735
            End
            Begin VB.CommandButton cmdCFOP_serv 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   12720
               Picture         =   "frmCompras_Pedido.frx":11A83A
               Style           =   1  'Graphical
               TabIndex        =   173
               ToolTipText     =   "Localizar CFOP de entrada."
               Top             =   990
               Width           =   315
            End
            Begin VB.TextBox txtCFOP_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   7650
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   171
               TabStop         =   0   'False
               ToolTipText     =   "Natureza da operação de entrada."
               Top             =   1005
               Width           =   1065
            End
            Begin VB.CheckBox chk_CFOP_serv 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7115
               TabIndex        =   169
               Top             =   780
               Width           =   195
            End
            Begin VB.TextBox txt_ID_CFOP_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   7115
               Locked          =   -1  'True
               TabIndex        =   170
               TabStop         =   0   'False
               ToolTipText     =   "ID da CFOP."
               Top             =   1005
               Width           =   525
            End
            Begin VB.CommandButton cmdLimpar_CFOP_serv 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   13050
               Picture         =   "frmCompras_Pedido.frx":11A93C
               Style           =   1  'Graphical
               TabIndex        =   174
               ToolTipText     =   "Limpar natureza de operação de entrada."
               Top             =   990
               Width           =   315
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2970
               Picture         =   "frmCompras_Pedido.frx":11AA7A
               Style           =   1  'Graphical
               TabIndex        =   162
               ToolTipText     =   "Visualizar arquivo."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox Txt_vlr_unit_ultima_compra_serv 
               Alignment       =   1  'Right Justify
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
               Left            =   13370
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   182
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário da última compra."
               Top             =   2385
               Width           =   1575
            End
            Begin VB.TextBox txtNatureza_operacao_serv 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   8730
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   172
               TabStop         =   0   'False
               ToolTipText     =   "Descrição da natureza da operação de entrada."
               Top             =   1005
               Width           =   3975
            End
            Begin DrawSuite2022.USButton imgCalendario2 
               Height          =   315
               Left            =   14640
               TabIndex        =   415
               Top             =   990
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":11B03C
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
               ShowFocusRect   =   0   'False
               Theme           =   3
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Detalhe"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   119
               Left            =   9125
               TabIndex        =   379
               Top             =   2190
               Width           =   555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "OP"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   118
               Left            =   11172
               TabIndex        =   378
               Top             =   2190
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "OS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   117
               Left            =   12652
               TabIndex        =   377
               Top             =   2190
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Observação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   116
               Left            =   11127
               TabIndex        =   376
               Top             =   1410
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   115
               Left            =   13865
               TabIndex        =   375
               Top             =   810
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Natureza de operação (Entrada)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   114
               Left            =   9547
               TabIndex        =   374
               Top             =   810
               Width           =   2340
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   113
               Left            =   7380
               TabIndex        =   373
               Top             =   810
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "CFOP"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   112
               Left            =   7980
               TabIndex        =   372
               Top             =   810
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
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
               Index           =   111
               Left            =   8835
               TabIndex        =   371
               Top             =   180
               Width           =   555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descrição comercial"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   36
               Left            =   3465
               TabIndex        =   231
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. unit. c/ desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   41
               Left            =   6822
               TabIndex        =   230
               Top             =   2820
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Código referência"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   33
               Left            =   4470
               TabIndex        =   229
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Família"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   37
               Left            =   4140
               TabIndex        =   228
               Top             =   2190
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   38
               Left            =   255
               TabIndex        =   227
               Top             =   2820
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor unitário"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   40
               Left            =   1980
               TabIndex        =   226
               Top             =   2820
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Código interno"
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
               Index           =   32
               Left            =   622
               TabIndex        =   225
               Top             =   180
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantidade"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   42
               Left            =   8700
               TabIndex        =   224
               Top             =   2820
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   45
               Left            =   13640
               TabIndex        =   223
               Top             =   2820
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "% ISSQN"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   43
               Left            =   10155
               TabIndex        =   222
               Top             =   2820
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor ISSQN"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   44
               Left            =   11662
               TabIndex        =   221
               Top             =   2820
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   39
               Left            =   955
               TabIndex        =   220
               Top             =   2820
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   35
               Left            =   3292
               TabIndex        =   219
               Top             =   810
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. última compra"
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
               Index           =   58
               Left            =   13385
               TabIndex        =   218
               Top             =   2190
               Width           =   1545
            End
         End
         Begin MSComctlLib.ListView Lista_custoServ 
            Height          =   5700
            Left            =   -74940
            TabIndex        =   201
            Top             =   945
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   10054
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   18600
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Percentual"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "ID_CC"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView ListaServ 
            Height          =   3735
            Left            =   60
            TabIndex        =   195
            Top             =   3795
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   6588
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
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
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6200
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Valor unit."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Desc. (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Vlr. desc."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Valor unit. c/ desc."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "ISSQN (%)"
               Object.Width           =   1640
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Vlr. ISSQN"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Valor total"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Status"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_empenhos_serv 
            Height          =   6315
            Left            =   -74940
            TabIndex        =   205
            Top             =   330
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   11139
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
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
            NumItems        =   17
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Cód. cart."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "N"
               Text            =   "Ped. int./SPR"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   2297
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Cod. ref."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   2914
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Qtde. rec."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Object.Tag             =   "T"
               Text            =   "Ped. cliente"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Object.Tag             =   "T"
               Text            =   "N. item"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar5 
         Height          =   975
         Left            =   -74925
         TabIndex        =   239
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   14
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Novo (Insert)"
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
         ButtonWidth1    =   33
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
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
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
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
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
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
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   47
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   220
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Status"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Alterar status do serviço (F9)"
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   268
         ButtonTop7      =   2
         ButtonWidth7    =   39
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Alterações"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Cadastrar alterações."
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   309
         ButtonTop8      =   2
         ButtonWidth8    =   59
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Centro de custo"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Copiar centro de custo."
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   370
         ButtonTop9      =   2
         ButtonWidth9    =   85
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Solicitação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Detalhes solicitação"
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   457
         ButtonTop10     =   2
         ButtonWidth10   =   58
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonAlignment11=   2
         ButtonType11    =   1
         ButtonStyle11   =   -1
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   -1
         ButtonLeft11    =   517
         ButtonTop11     =   4
         ButtonWidth11   =   2
         ButtonHeight11  =   54
         ButtonCaption12 =   "Ajuda"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Ajuda (F1)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   521
         ButtonTop12     =   2
         ButtonWidth12   =   36
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Sair"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Sair (Esc)"
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   559
         ButtonTop13     =   2
         ButtonWidth13   =   26
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   587
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
         Begin DrawSuite2022.USImageList USImageList5 
            Left            =   13560
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Pedido.frx":121E09
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8685
         Left            =   75
         TabIndex        =   240
         Top             =   1350
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15319
         _Version        =   393216
         TabHeight       =   520
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Dados do produto"
         TabPicture(0)   =   "frmCompras_Pedido.frx":12979A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Listprod"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtcodproduto"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtidcarteira"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TXTIDLista"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame1(12)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtQuantidade_PC"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Centro de custo"
         TabPicture(1)   =   "frmCompras_Pedido.frx":1297B6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista_custo"
         Tab(1).Control(1)=   "Frame1(5)"
         Tab(1).Control(2)=   "txtIDCentro"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame1(13)"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Empenhos"
         TabPicture(2)   =   "frmCompras_Pedido.frx":1297D2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(6)"
         Tab(2).Control(1)=   "Lista_empenhos"
         Tab(2).ControlCount=   2
         Begin VB.TextBox txtQuantidade_PC 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2730
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   395
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade em peça."
            Top             =   4890
            Visible         =   0   'False
            Width           =   995
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Index           =   6
            Left            =   -74940
            TabIndex        =   297
            Top             =   6660
            Width           =   15135
            Begin VB.TextBox Txt_qtde_total_comprada 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   9660
               Locked          =   -1  'True
               TabIndex        =   156
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade total comprada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_emp 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   11490
               Locked          =   -1  'True
               TabIndex        =   157
               TabStop         =   0   'False
               ToolTipText     =   "Quatidade total empenhada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_disp 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   13380
               Locked          =   -1  'True
               TabIndex        =   158
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade disponível."
               Top             =   420
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. disponível"
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
               Index           =   120
               Left            =   13492
               TabIndex        =   382
               Top             =   210
               Width           =   1350
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. empenhada"
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
               Index           =   51
               Left            =   11527
               TabIndex        =   381
               Top             =   210
               Width           =   1500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. comprada"
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
               Index           =   5
               Left            =   9772
               TabIndex        =   380
               Top             =   210
               Width           =   1350
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "-                                       ="
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
               Index           =   47
               Left            =   11310
               TabIndex        =   298
               Top             =   480
               Width           =   1965
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   3480
            Index           =   12
            Left            =   30
            TabIndex        =   249
            Top             =   330
            Width           =   15195
            Begin VB.ComboBox cmbReferencia 
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
               ItemData        =   "frmCompras_Pedido.frx":1297EE
               Left            =   3390
               List            =   "frmCompras_Pedido.frx":1297F0
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   101
               ToolTipText     =   "Código de referencia."
               Top             =   390
               Width           =   3435
            End
            Begin VB.TextBox txt_ID_CFOP_prod 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   6090
               Locked          =   -1  'True
               TabIndex        =   107
               TabStop         =   0   'False
               ToolTipText     =   "ID da CFOP."
               Top             =   1005
               Width           =   525
            End
            Begin VB.TextBox txtReferencia 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   3390
               MaxLength       =   255
               TabIndex        =   100
               ToolTipText     =   "Código de referência."
               Top             =   390
               Visible         =   0   'False
               Width           =   3435
            End
            Begin VB.Frame framePrazo 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   12285
               TabIndex        =   251
               Top             =   1000
               Width           =   945
               Begin MSMask.MaskEdBox txtprazo_item 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   112
                  ToolTipText     =   "Prazo de entrega."
                  Top             =   0
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   0
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
            End
            Begin VB.TextBox TxtVlrTotal 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   13410
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   141
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   3015
               Width           =   1515
            End
            Begin VB.TextBox TxtVlrIcms 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   9225
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   137
               TabStop         =   0   'False
               ToolTipText     =   "Valor de ICMS."
               Top             =   3015
               Width           =   1050
            End
            Begin VB.TextBox TxtvlrIpi 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   7575
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   135
               TabStop         =   0   'False
               ToolTipText     =   "Valor de IPI."
               Top             =   3015
               Width           =   1035
            End
            Begin VB.TextBox txtIcms 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   8625
               MaxLength       =   50
               TabIndex        =   136
               ToolTipText     =   "Valor de % do  ICMS."
               Top             =   3015
               Width           =   585
            End
            Begin VB.TextBox txtQuantidade 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   5130
               MaxLength       =   50
               TabIndex        =   132
               ToolTipText     =   "Quantidade da unidade comercial."
               Top             =   3015
               Width           =   1005
            End
            Begin VB.TextBox txtNomenclatura 
               Alignment       =   2  'Center
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
               Left            =   180
               MaxLength       =   50
               TabIndex        =   96
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   2115
            End
            Begin VB.CommandButton cmdfiltrar 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2310
               Picture         =   "frmCompras_Pedido.frx":1297F2
               Style           =   1  'Graphical
               TabIndex        =   97
               ToolTipText     =   "Filtrar por código interno."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtvalorunitario 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   180
               MaxLength       =   50
               TabIndex        =   128
               ToolTipText     =   "Valor unitário."
               Top             =   3015
               Width           =   1065
            End
            Begin VB.TextBox txtipi 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   7005
               MaxLength       =   50
               TabIndex        =   134
               ToolTipText     =   "Valor de % do IPI."
               Top             =   3015
               Width           =   555
            End
            Begin VB.ComboBox cmbun 
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
               ItemData        =   "frmCompras_Pedido.frx":129C0D
               Left            =   4350
               List            =   "frmCompras_Pedido.frx":129C0F
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   117
               TabStop         =   0   'False
               ToolTipText     =   "Unidade de estoque."
               Top             =   2385
               Width           =   735
            End
            Begin VB.ComboBox cmbfamilia 
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
               Left            =   180
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   116
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   2385
               Width           =   4155
            End
            Begin VB.TextBox txtEspecificacoes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Height          =   315
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   105
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   1005
               Width           =   5895
            End
            Begin VB.TextBox txtdetalheitem 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   9060
               MaxLength       =   50
               TabIndex        =   124
               ToolTipText     =   "Detalhe."
               Top             =   2385
               Width           =   2225
            End
            Begin VB.TextBox txtObs 
               Alignment       =   2  'Center
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
               Height          =   465
               Left            =   8180
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   115
               ToolTipText     =   "Observações."
               Top             =   1620
               Width           =   6765
            End
            Begin VB.TextBox txtvalorunitariodesc 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   3720
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   131
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   3015
               Width           =   1395
            End
            Begin VB.TextBox txtvalordesconto 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   2340
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   130
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   3015
               Width           =   1365
            End
            Begin VB.TextBox txtdesconto 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   1260
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   129
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   3015
               Width           =   1065
            End
            Begin VB.CommandButton CmdEscolher_item 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2640
               Picture         =   "frmCompras_Pedido.frx":129C11
               Style           =   1  'Graphical
               TabIndex        =   98
               ToolTipText     =   "Localizar produtos."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtstatus_item 
               Alignment       =   2  'Center
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
               Left            =   6840
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   102
               TabStop         =   0   'False
               ToolTipText     =   "Status."
               Top             =   390
               Width           =   4545
            End
            Begin VB.CheckBox chkRemessa 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Remessa"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   13590
               TabIndex        =   113
               Top             =   1005
               Width           =   1095
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto"
               ForeColor       =   &H00000000&
               Height          =   615
               Index           =   14
               Left            =   11640
               TabIndex        =   250
               Top             =   180
               Width           =   3285
               Begin VB.CheckBox chkAuto 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. automático ?"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   180
                  TabIndex        =   103
                  Top             =   270
                  Width           =   1605
               End
               Begin VB.CheckBox chkManual 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. manual ?"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   1890
                  TabIndex        =   104
                  Top             =   270
                  Width           =   1335
               End
            End
            Begin VB.TextBox txtOrdem 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   11295
               TabIndex        =   125
               ToolTipText     =   "Número da ordem de produção."
               Top             =   2385
               Width           =   1035
            End
            Begin VB.CheckBox Chk_desc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desc. (%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1275
               TabIndex        =   142
               Top             =   2820
               Width           =   1035
            End
            Begin VB.CheckBox Chk_valor_desc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Vlr. do desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2415
               TabIndex        =   143
               Top             =   2820
               Width           =   1215
            End
            Begin VB.TextBox txtDescricao_comercial 
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
               Height          =   465
               Left            =   180
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   114
               ToolTipText     =   "Descrição comercial."
               Top             =   1620
               Width           =   7965
            End
            Begin VB.ComboBox Cmb_un_com 
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
               ItemData        =   "frmCompras_Pedido.frx":129D13
               Left            =   5085
               List            =   "frmCompras_Pedido.frx":129D15
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   118
               ToolTipText     =   "Unidade comercial."
               Top             =   2385
               Width           =   735
            End
            Begin VB.ComboBox Cmb_OS 
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
               ItemData        =   "frmCompras_Pedido.frx":129D17
               Left            =   12345
               List            =   "frmCompras_Pedido.frx":129D19
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   126
               ToolTipText     =   "Número da OS."
               Top             =   2385
               Width           =   1185
            End
            Begin VB.TextBox txtQuantidade_est 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   6150
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   133
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade da unidade de estoque."
               Top             =   3015
               Width           =   840
            End
            Begin VB.CheckBox Chk_CFOP_prod 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6090
               TabIndex        =   106
               Top             =   780
               Width           =   195
            End
            Begin VB.CommandButton cmdCFOP_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   11600
               Picture         =   "frmCompras_Pedido.frx":129D1B
               Style           =   1  'Graphical
               TabIndex        =   110
               ToolTipText     =   "Localizar CFOP de entrada."
               Top             =   1005
               Width           =   315
            End
            Begin VB.TextBox txt_Natureza_operacao_prod 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   7710
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   109
               TabStop         =   0   'False
               ToolTipText     =   "Descrição da natureza da operação de entrada."
               Top             =   1005
               Width           =   3855
            End
            Begin VB.CommandButton cmdCF 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7365
               Picture         =   "frmCompras_Pedido.frx":129E1D
               Style           =   1  'Graphical
               TabIndex        =   121
               ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
               Top             =   2385
               Width           =   315
            End
            Begin VB.TextBox Txt_ID_CF 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   5820
               Locked          =   -1  'True
               TabIndex        =   119
               TabStop         =   0   'False
               ToolTipText     =   "ID da NCM."
               Top             =   2385
               Width           =   525
            End
            Begin VB.ComboBox Cmb_CST_ICMS 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCompras_Pedido.frx":129F1F
               Left            =   8115
               List            =   "frmCompras_Pedido.frx":129FDA
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   123
               ToolTipText     =   "Situação tributária ICMS."
               Top             =   2370
               Width           =   930
            End
            Begin VB.CommandButton cmdLimpar_NCM 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   7695
               Picture         =   "frmCompras_Pedido.frx":12A12A
               Style           =   1  'Graphical
               TabIndex        =   122
               ToolTipText     =   "Limpar classificação fiscal."
               Top             =   2385
               Width           =   315
            End
            Begin VB.CommandButton cmdLimpar_CFOP_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11940
               Picture         =   "frmCompras_Pedido.frx":12A268
               Style           =   1  'Graphical
               TabIndex        =   111
               ToolTipText     =   "Limpar natureza de operação de entrada."
               Top             =   1005
               Width           =   315
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2970
               Picture         =   "frmCompras_Pedido.frx":12A3A6
               Style           =   1  'Graphical
               TabIndex        =   99
               ToolTipText     =   "Visualizar arquivo."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtFrete 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   10290
               MaxLength       =   50
               TabIndex        =   138
               ToolTipText     =   "Valor do frete."
               Top             =   3015
               Width           =   885
            End
            Begin VB.TextBox txtAcessorias 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   12210
               MaxLength       =   50
               TabIndex        =   140
               ToolTipText     =   "Valor das despesas acessórias."
               Top             =   3015
               Width           =   1205
            End
            Begin VB.TextBox txtSeguro 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   11190
               MaxLength       =   50
               TabIndex        =   139
               ToolTipText     =   "Valor do seguro."
               Top             =   3015
               Width           =   1005
            End
            Begin VB.CheckBox ChkFrete_IPI 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Tem IPI no frete"
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
               Height          =   345
               Left            =   13590
               TabIndex        =   144
               Top             =   1200
               Width           =   1365
            End
            Begin VB.TextBox Txt_vlr_unit_ultima_compra_prod 
               Alignment       =   1  'Right Justify
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
               Left            =   13550
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   127
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário da última compra."
               Top             =   2385
               Width           =   1395
            End
            Begin VB.TextBox Txt_CF 
               Alignment       =   2  'Center
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
               Height          =   315
               Left            =   6360
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   120
               TabStop         =   0   'False
               ToolTipText     =   "Classificação fiscal."
               Top             =   2385
               Width           =   975
            End
            Begin VB.TextBox txtCFOP_prod 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   6630
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   108
               TabStop         =   0   'False
               ToolTipText     =   "Natureza da operação de entrada."
               Top             =   1005
               Width           =   1065
            End
            Begin DrawSuite2022.USButton imgCalendario 
               Height          =   315
               Left            =   13200
               TabIndex        =   414
               Top             =   990
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":12A968
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
               ShowFocusRect   =   0   'False
               Theme           =   3
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "OS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   108
               Left            =   12832
               TabIndex        =   368
               Top             =   2190
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "OP"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   107
               Left            =   11707
               TabIndex        =   367
               Top             =   2190
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Detalhe"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   106
               Left            =   9895
               TabIndex        =   366
               Top             =   2190
               Width           =   555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "CST ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   105
               Left            =   8228
               TabIndex        =   365
               Top             =   2190
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "NCM"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   104
               Left            =   6682
               TabIndex        =   364
               Top             =   2190
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   103
               Left            =   6000
               TabIndex        =   363
               Top             =   2190
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   102
               Left            =   5130
               TabIndex        =   362
               Top             =   2190
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   101
               Left            =   4425
               TabIndex        =   361
               Top             =   2190
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Observação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   100
               Left            =   11127
               TabIndex        =   360
               Top             =   1410
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   99
               Left            =   12630
               TabIndex        =   359
               Top             =   810
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Natureza de operação (Entrada)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   98
               Left            =   8467
               TabIndex        =   358
               Top             =   810
               Width           =   2340
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "CFOP"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   97
               Left            =   6960
               TabIndex        =   357
               Top             =   810
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   96
               Left            =   6360
               TabIndex        =   356
               Top             =   810
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
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
               Index           =   95
               Left            =   8835
               TabIndex        =   355
               Top             =   180
               Width           =   555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   30
               Left            =   13815
               TabIndex        =   269
               Top             =   2820
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   29
               Left            =   9427
               TabIndex        =   268
               Top             =   2820
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. IPI"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   27
               Left            =   7920
               TabIndex        =   267
               Top             =   2820
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "% ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   28
               Left            =   8625
               TabIndex        =   266
               Top             =   2820
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtd. com."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   25
               Left            =   5265
               TabIndex        =   265
               Top             =   2820
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Código interno"
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
               Index           =   16
               Left            =   622
               TabIndex        =   264
               Top             =   180
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. unitário"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   22
               Left            =   225
               TabIndex        =   263
               Top             =   2820
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "% IPI"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   7095
               TabIndex        =   262
               Top             =   2820
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Código referência"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   17
               Left            =   4470
               TabIndex        =   261
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descrição comercial"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   20
               Left            =   3540
               TabIndex        =   260
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. unit. c/ desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   23
               Left            =   3802
               TabIndex        =   259
               Top             =   2820
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Qt.  est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   26
               Left            =   6225
               TabIndex        =   258
               Top             =   2820
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   19
               Left            =   3090
               TabIndex        =   257
               Top             =   810
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Família"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   21
               Left            =   2017
               TabIndex        =   256
               Top             =   2190
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "vlr. Frete"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   24
               Left            =   10395
               TabIndex        =   255
               Top             =   2820
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "vlr. Desp. aces."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   52
               Left            =   12225
               TabIndex        =   254
               Top             =   2820
               Width           =   1140
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "vlr.Seguro"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   54
               Left            =   11310
               TabIndex        =   253
               Top             =   2820
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. últ. compra"
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
               Index           =   57
               Left            =   13610
               TabIndex        =   252
               Top             =   2190
               Width           =   1275
            End
         End
         Begin VB.TextBox TXTIDLista 
            Height          =   315
            Left            =   420
            TabIndex        =   248
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4905
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtidcarteira 
            Enabled         =   0   'False
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
            Left            =   1980
            MouseIcon       =   "frmCompras_Pedido.frx":131735
            MousePointer    =   99  'Custom
            TabIndex        =   247
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4905
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtcodproduto 
            Height          =   315
            Left            =   1200
            TabIndex        =   246
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4905
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   600
            Index           =   13
            Left            =   -74940
            TabIndex        =   244
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtValorCentro 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   11295
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   148
               TabStop         =   0   'False
               ToolTipText     =   "Valor."
               Top             =   180
               Width           =   1155
            End
            Begin VB.TextBox txtPercentualCentro 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   13785
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   150
               TabStop         =   0   'False
               ToolTipText     =   "Percentual."
               Top             =   180
               Width           =   1155
            End
            Begin VB.CheckBox chkValor 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor :"
               Height          =   255
               Left            =   10530
               TabIndex        =   147
               Top             =   180
               Width           =   765
            End
            Begin VB.CheckBox chkPercentual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Percentual :"
               Height          =   255
               Left            =   12600
               TabIndex        =   149
               Top             =   180
               Width           =   1185
            End
            Begin VB.ComboBox Cmb_centro 
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
               Left            =   1500
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   146
               ToolTipText     =   "Centro de custo."
               Top             =   180
               Width           =   8910
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Centro de custo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   31
               Left            =   180
               TabIndex        =   245
               Top             =   180
               Width           =   1260
            End
         End
         Begin VB.TextBox txtIDCentro 
            Enabled         =   0   'False
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
            Left            =   -74370
            MouseIcon       =   "frmCompras_Pedido.frx":131A3F
            MousePointer    =   99  'Custom
            TabIndex        =   243
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1650
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   855
            Index           =   5
            Left            =   -74940
            TabIndex        =   241
            Top             =   6660
            Width           =   15135
            Begin VB.TextBox txtVlrTotal_centro 
               Alignment       =   1  'Right Justify
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
               Left            =   10200
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   152
               TabStop         =   0   'False
               ToolTipText     =   "Valor total do produto/item."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtSaldoCentro 
               Alignment       =   1  'Right Justify
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
               Left            =   13380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   154
               TabStop         =   0   'False
               ToolTipText     =   "Saldo."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtTotalCentro 
               Alignment       =   1  'Right Justify
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
               Left            =   11790
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   153
               TabStop         =   0   'False
               ToolTipText     =   "Valor total centro de custo."
               Top             =   420
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo"
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
               Index           =   110
               Left            =   13935
               TabIndex        =   370
               Top             =   210
               Width           =   465
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total centro"
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
               Index           =   109
               Left            =   12060
               TabIndex        =   369
               Top             =   210
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total"
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
               Index           =   50
               Left            =   10545
               TabIndex        =   242
               Top             =   210
               Width           =   885
            End
         End
         Begin MSComctlLib.ListView Listprod 
            Height          =   3705
            Left            =   30
            TabIndex        =   145
            Top             =   3825
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   6535
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
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
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   4313
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde. est."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Qtde. com."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Valor unit."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Desc. (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Vlr. desc."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Valor unit. c/ desc."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "IPI (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "ICMS (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Vlr. IPI"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "Valor total"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_custo 
            Height          =   5700
            Left            =   -74940
            TabIndex        =   151
            Top             =   945
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   10054
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   18600
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Percentual"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "ID_CC"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_empenhos 
            Height          =   6315
            Left            =   -74940
            TabIndex        =   155
            Top             =   330
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   11139
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
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
            NumItems        =   17
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Cód. cart."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "N"
               Text            =   "Ped. int./SPR"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   2297
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Cod. ref."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   2914
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Qtde. rec."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Object.Tag             =   "T"
               Text            =   "Ped. cliente"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Object.Tag             =   "T"
               Text            =   "N. item"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   1035
         Left            =   75
         TabIndex        =   270
         Top             =   330
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1826
         ButtonCount     =   14
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDownRight1=   10802943
         GradientColorDownRight2=   7979263
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOverRight1=   14417407
         GradientColorOverRight2=   12317439
         IsStrech        =   -1  'True
         RightColor1     =   14737632
         RightColor2     =   16777215
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Novo (Insert)"
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
         ButtonWidth1    =   33
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
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
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
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
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
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
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   47
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   220
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Calculadora"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Abrir calculadora para cálculo de peso (F8)"
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   268
         ButtonTop7      =   2
         ButtonWidth7    =   64
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Status"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Alterar status do produto (F9)"
         ButtonKey8      =   "9"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   334
         ButtonTop8      =   2
         ButtonWidth8    =   39
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Alterações"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Cadastrar alterações."
         ButtonKey9      =   "10"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   375
         ButtonTop9      =   2
         ButtonWidth9    =   59
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Centro de custo"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Copiar centro de custo."
         ButtonKey10     =   "11"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   436
         ButtonTop10     =   2
         ButtonWidth10   =   85
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Necess./Solici."
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Lista de necessidades/solicitações"
         ButtonKey11     =   "12"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   523
         ButtonTop11     =   2
         ButtonWidth11   =   77
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonAlignment12=   2
         ButtonType12    =   1
         ButtonStyle12   =   -1
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   -1
         ButtonLeft12    =   602
         ButtonTop12     =   4
         ButtonWidth12   =   2
         ButtonHeight12  =   58
         ButtonCaption13 =   "Ajuda"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Ajuda (F1)"
         ButtonKey13     =   "14"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   606
         ButtonTop13     =   2
         ButtonWidth13   =   36
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Sair"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Sair (Esc)"
         ButtonKey14     =   "15"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   644
         ButtonTop14     =   2
         ButtonWidth14   =   26
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   12150
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Pedido.frx":131D49
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   281
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   9
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
         ButtonWidth1    =   38
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
         ButtonLeft2     =   42
         ButtonTop2      =   2
         ButtonWidth2    =   51
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Anterior"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Registro anterior."
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
         ButtonLeft3     =   95
         ButtonTop3      =   2
         ButtonWidth3    =   47
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Próximo"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Próximo registro."
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
         ButtonLeft4     =   144
         ButtonTop4      =   2
         ButtonWidth4    =   46
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Financeiro"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Enviar para o financeiro (F7)"
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
         ButtonLeft5     =   192
         ButtonTop5      =   2
         ButtonWidth5    =   57
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonAlignment6=   2
         ButtonType6     =   1
         ButtonStyle6    =   -1
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   -1
         ButtonLeft6     =   251
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   54
         ButtonCaption7  =   "Ajuda"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Ajuda (F1)"
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   255
         ButtonTop7      =   2
         ButtonWidth7    =   36
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Sair"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Sair (Esc)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   293
         ButtonTop8      =   2
         ButtonWidth8    =   26
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   5
         ButtonLeft9     =   321
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         ButtonUseMaskColor9=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   13620
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Pedido.frx":139EB1
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView listapedido 
         Height          =   3945
         Left            =   -74910
         TabIndex        =   73
         Top             =   4320
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   6959
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Cotação"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   11704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Validado"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Aprovado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Valor Pedido"
            Object.Width           =   2540
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   288
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   1720
         ButtonCount     =   17
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Novo (Insert)"
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
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Filtrar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Filtrar (F2)"
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
         ButtonWidth2    =   36
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Salvar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Salvar (F3)"
         ButtonKey3      =   "3"
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
         ButtonLeft3     =   78
         ButtonTop3      =   2
         ButtonWidth3    =   44
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Status"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Alterar status do pedido (F4)"
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
         ButtonLeft4     =   124
         ButtonTop4      =   2
         ButtonWidth4    =   45
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Relatório"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Relatório (F5)"
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
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   60
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Anterior"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Registro anterior."
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   233
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Próximo"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Próximo registro."
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   290
         ButtonTop7      =   2
         ButtonWidth7    =   55
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar (F7)"
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   347
         ButtonTop8      =   2
         ButtonWidth8    =   39
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Validação"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Validar/Cancelar validação (F9)"
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   388
         ButtonTop9      =   2
         ButtonWidth9    =   53
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Aprovação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Aprovar/cancelar aprovação (F10)"
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   443
         ButtonTop10     =   2
         ButtonWidth10   =   60
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Enviar Email"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Enviar pedido por e-mail (F11]"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   505
         ButtonTop11     =   2
         ButtonWidth11   =   65
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Exportar"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Exportar para excel (F12)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   572
         ButtonTop12     =   2
         ButtonWidth12   =   50
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Atualizar"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Utilizado pelo administrador do sistema."
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   624
         ButtonTop13     =   2
         ButtonWidth13   =   50
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonAlignment14=   2
         ButtonType14    =   1
         ButtonStyle14   =   -1
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   -1
         ButtonLeft14    =   676
         ButtonTop14     =   4
         ButtonWidth14   =   2
         ButtonHeight14  =   54
         ButtonCaption15 =   "Ajuda"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Ajuda (F1)"
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   680
         ButtonTop15     =   2
         ButtonWidth15   =   41
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonCaption16 =   "Sair"
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonToolTipText16=   "Sair (Esc)"
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft16    =   723
         ButtonTop16     =   2
         ButtonWidth16   =   30
         ButtonHeight16  =   21
         ButtonUseMaskColor16=   0   'False
         ButtonEnabled17 =   0   'False
         ButtonIconSize17=   32
         ButtonKey17     =   "17"
         ButtonAlignment17=   2
         BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState17   =   5
         ButtonLeft17    =   755
         ButtonTop17     =   2
         ButtonWidth17   =   24
         ButtonHeight17  =   24
         ButtonUseMaskColor17=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13830
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Pedido.frx":13E54B
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74880
         TabIndex        =   304
         Top             =   630
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   6
         GradientColor1  =   16777215
         GradientColor2  =   14737632
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDownRight1=   10802943
         GradientColorDownRight2=   7979263
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOverRight1=   14417407
         GradientColorOverRight2=   12317439
         IsStrech        =   -1  'True
         RightColor1     =   14737632
         RightColor2     =   16777215
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
         ButtonCaption2  =   "Gerar ped."
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Gerar pedido (F3)"
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
         ButtonWidth2    =   67
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
         ButtonLeft3     =   115
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
         ButtonLeft4     =   119
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
         ButtonLeft5     =   162
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   194
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   5520
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Pedido.frx":1488DE
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   389
         Top             =   9720
         Width           =   15105
         _ExtentX        =   26644
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
      Begin TabDlg.SSTab SSTab4 
         Height          =   9720
         Left            =   -74940
         TabIndex        =   305
         Top             =   300
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   17145
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Necessidade"
         TabPicture(0)   =   "frmCompras_Pedido.frx":14B42B
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ListaFabricante"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame10"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame1(23)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame15"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "ListaNecessidade"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Solicitação"
         TabPicture(1)   =   "frmCompras_Pedido.frx":14B447
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista_solicitados"
         Tab(1).Control(1)=   "Frame1(2)"
         Tab(1).Control(2)=   "Frame7"
         Tab(1).ControlCount=   3
         Begin MSComctlLib.ListView ListaNecessidade 
            Height          =   5910
            Left            =   75
            TabIndex        =   11
            Top             =   2760
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   10425
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   14578
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Un."
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Necessidade"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00E0E0E0&
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
            Height          =   675
            Left            =   4770
            TabIndex        =   312
            Top             =   1320
            Width           =   10485
            Begin VB.OptionButton Opt_vendas 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Comprar para  vendas"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   3000
               TabIndex        =   2
               Top             =   270
               Width           =   2445
            End
            Begin VB.OptionButton Opt_PCP 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Comprar para produção"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   150
               TabIndex        =   1
               Top             =   270
               Value           =   -1  'True
               Width           =   3045
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   -74925
            TabIndex        =   317
            Top             =   8790
            Width           =   15105
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
               Index           =   2
               Left            =   3780
               TabIndex        =   28
               Text            =   "30"
               ToolTipText     =   "Número de registros por página."
               Top             =   180
               Width           =   555
            End
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
               Index           =   2
               Left            =   9540
               TabIndex        =   29
               ToolTipText     =   "Número da página."
               Top             =   180
               Width           =   555
            End
            Begin DrawSuite2022.USButton cmdPagProx 
               Height          =   315
               Index           =   2
               Left            =   11760
               TabIndex        =   33
               ToolTipText     =   "Próxima página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":14B463
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagAnt 
               Height          =   315
               Index           =   2
               Left            =   11220
               TabIndex        =   32
               ToolTipText     =   "Página anterior."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":14EC07
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagIr 
               Height          =   315
               Index           =   2
               Left            =   10110
               TabIndex        =   30
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   556
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagPrim 
               Height          =   315
               Index           =   2
               Left            =   10680
               TabIndex        =   31
               ToolTipText     =   "Primeira página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":152710
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagUlt 
               Height          =   315
               Index           =   2
               Left            =   12300
               TabIndex        =   34
               ToolTipText     =   "Última página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":1567FF
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "registros por página"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   127
               Left            =   4410
               TabIndex        =   391
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Carregar"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   129
               Left            =   3090
               TabIndex        =   320
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lblPaginas 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Página: 0 de: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   13050
               TabIndex        =   319
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lblRegistros 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de registros: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   318
               Top             =   240
               Width           =   1275
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Index           =   23
            Left            =   75
            TabIndex        =   313
            Top             =   1920
            Width           =   15165
            Begin VB.Frame Frame1 
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
               Height          =   510
               Index           =   17
               Left            =   3240
               TabIndex        =   393
               Top             =   195
               Width           =   4785
               Begin VB.OptionButton optIgual_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Igual"
                  Height          =   255
                  Left            =   3930
                  TabIndex        =   10
                  Top             =   180
                  Width           =   705
               End
               Begin VB.OptionButton Optmeio_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Meio"
                  Height          =   255
                  Left            =   1470
                  TabIndex        =   8
                  Top             =   180
                  Width           =   1275
               End
               Begin VB.OptionButton optInicio_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Início"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   7
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   1275
               End
               Begin VB.OptionButton Optfim_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Fim"
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   9
                  Top             =   180
                  Width           =   1155
               End
            End
            Begin VB.ComboBox cmbfiltrarpor_necess 
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
               Height          =   330
               ItemData        =   "frmCompras_Pedido.frx":15A08B
               Left            =   180
               List            =   "frmCompras_Pedido.frx":15A09E
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   3
               ToolTipText     =   "Opções para filtro."
               Top             =   390
               Width           =   2955
            End
            Begin VB.ComboBox Cmb_filtrar 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmCompras_Pedido.frx":15A0E9
               Left            =   12630
               List            =   "frmCompras_Pedido.frx":15A0F3
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               ToolTipText     =   "Tipo de necessidade."
               Top             =   390
               Width           =   2295
            End
            Begin VB.TextBox txtTexto_necess 
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
               Height          =   315
               Left            =   8130
               TabIndex        =   4
               ToolTipText     =   "Texto para pesquisa."
               Top             =   390
               Width           =   4485
            End
            Begin VB.ComboBox cmbTexto_necess 
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
               ItemData        =   "frmCompras_Pedido.frx":15A11D
               Left            =   8130
               List            =   "frmCompras_Pedido.frx":15A11F
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   5
               ToolTipText     =   "Texto para pesquisa."
               Top             =   390
               Visible         =   0   'False
               Width           =   4485
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texto para pesquisa"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   60
               Left            =   9630
               TabIndex        =   316
               Top             =   180
               Width           =   1485
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filtrar por"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   59
               Left            =   1230
               TabIndex        =   315
               Top             =   180
               Width           =   705
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de necessidade"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12915
               TabIndex        =   314
               Top             =   180
               Width           =   1455
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   75
            TabIndex        =   308
            Top             =   8790
            Width           =   15105
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
               Index           =   1
               Left            =   9540
               TabIndex        =   13
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
               Index           =   1
               Left            =   3780
               TabIndex        =   12
               Text            =   "30"
               ToolTipText     =   "Número de registros por página."
               Top             =   180
               Width           =   555
            End
            Begin DrawSuite2022.USButton cmdPagProx 
               Height          =   315
               Index           =   1
               Left            =   11760
               TabIndex        =   17
               ToolTipText     =   "Próxima página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":15A121
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagAnt 
               Height          =   315
               Index           =   1
               Left            =   11220
               TabIndex        =   16
               ToolTipText     =   "Página anterior."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":15D8C5
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagIr 
               Height          =   315
               Index           =   1
               Left            =   10110
               TabIndex        =   14
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   556
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagPrim 
               Height          =   315
               Index           =   1
               Left            =   10680
               TabIndex        =   15
               ToolTipText     =   "Primeira página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":1613CE
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin DrawSuite2022.USButton cmdPagUlt 
               Height          =   315
               Index           =   1
               Left            =   12300
               TabIndex        =   18
               ToolTipText     =   "Última página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmCompras_Pedido.frx":1654BD
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
               BorderColor     =   14404026
               BorderColorDown =   11632444
               BorderColorOver =   11632444
               GradientColor2  =   16777215
               GradientColor3  =   16777215
               GradientColorOver1=   16643560
               GradientColorOver2=   16576988
               GradientColorOver3=   16441780
               GradientColorOver4=   16178091
               GradientColorDown2=   16246986
               GradientColorDown3=   15189380
               GradientColorDown4=   14596208
               PicSizeH        =   19
               PicSizeW        =   19
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "registros por página"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   126
               Left            =   4410
               TabIndex        =   390
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label lblRegistros 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de registros: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   311
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label lblPaginas 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Página: 0 de: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   13050
               TabIndex        =   310
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Carregar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3090
               TabIndex        =   309
               Top             =   240
               Width           =   645
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Filtrar por"
            Height          =   675
            Index           =   2
            Left            =   -70230
            TabIndex        =   306
            Top             =   1320
            Width           =   10485
            Begin VB.Frame Frame1 
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
               Height          =   690
               Index           =   18
               Left            =   2850
               TabIndex        =   394
               Top             =   -10
               Width           =   3105
               Begin VB.OptionButton optFim_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Fim"
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   24
                  Top             =   270
                  Width           =   645
               End
               Begin VB.OptionButton optInicio_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Início"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   23
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton optMeio_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Meio"
                  Height          =   255
                  Left            =   870
                  TabIndex        =   25
                  Top             =   270
                  Width           =   705
               End
               Begin VB.OptionButton optIgual_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Igual"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   26
                  Top             =   270
                  Width           =   705
               End
            End
            Begin VB.ComboBox cmbfiltrarpor_sol 
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
               Height          =   330
               ItemData        =   "frmCompras_Pedido.frx":168D49
               Left            =   180
               List            =   "frmCompras_Pedido.frx":168D65
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   19
               ToolTipText     =   "Opções para filtro."
               Top             =   240
               Width           =   2595
            End
            Begin VB.TextBox txtTexto_sol 
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
               Height          =   315
               Left            =   6210
               TabIndex        =   20
               ToolTipText     =   "Texto para pesquisa."
               Top             =   240
               Width           =   4005
            End
            Begin MSComCtl2.DTPicker Txtprazo_sol 
               Height          =   315
               Left            =   6240
               TabIndex        =   22
               ToolTipText     =   "Texto para pesquisa."
               Top             =   240
               Visible         =   0   'False
               Width           =   4005
               _ExtentX        =   7064
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
               Format          =   184287233
               CurrentDate     =   39057
            End
            Begin VB.ComboBox cmbTexto_sol 
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
               ItemData        =   "frmCompras_Pedido.frx":168DD4
               Left            =   6240
               List            =   "frmCompras_Pedido.frx":168DD6
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   21
               ToolTipText     =   "Familia."
               Top             =   240
               Visible         =   0   'False
               Width           =   4005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texto para pesquisa"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   18
               Left            =   7170
               TabIndex        =   307
               Top             =   180
               Width           =   1485
            End
         End
         Begin MSComctlLib.ListView Lista_solicitados 
            Height          =   6765
            Left            =   -74925
            TabIndex        =   27
            Top             =   2010
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   11933
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Nº solicitação"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6729
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Un. est."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Un. com."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Quant. est."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Quant. com."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Detalhe"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Object.Tag             =   "D"
               Text            =   "Prazo entr."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "T"
               Text            =   "Obs."
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView ListaFabricante 
            Height          =   5910
            Left            =   10710
            TabIndex        =   413
            Top             =   2760
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   10425
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Part Number"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Fabricante"
               Object.Width           =   6068
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do pedido"
         ForeColor       =   &H00000000&
         Height          =   945
         Index           =   0
         Left            =   -74925
         TabIndex        =   289
         Top             =   1320
         Width           =   15225
         Begin VB.CheckBox Chk_email_enviado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Email enviado?"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   13650
            TabIndex        =   38
            Top             =   0
            Width           =   1545
         End
         Begin VB.TextBox txtresponsavel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   1770
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   495
            Width           =   2115
         End
         Begin VB.TextBox txtdata 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   495
            Width           =   735
         End
         Begin VB.TextBox txtpedido 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Número do pedido de compra."
            Top             =   495
            Width           =   855
         End
         Begin VB.TextBox txtData_aprovacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Data da aprovação."
            Top             =   495
            Width           =   855
         End
         Begin VB.TextBox txtResponsavel_aprovacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   8130
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela aprovação."
            Top             =   495
            Width           =   1545
         End
         Begin VB.TextBox txtDtValidacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "10/03/2020 14:12:30"
            ToolTipText     =   "Data e hora da validação."
            Top             =   495
            Width           =   1545
         End
         Begin VB.TextBox txtRespValidacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   5460
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   495
            Width           =   1785
         End
         Begin VB.TextBox txtstatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9690
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Status do pedido."
            Top             =   495
            Width           =   2205
         End
         Begin DrawSuite2022.USButton cmdstatus 
            Height          =   315
            Left            =   12240
            TabIndex        =   401
            Top             =   495
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":168DD8
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin DrawSuite2022.USButton btnStatusPedido 
            Height          =   315
            Left            =   11910
            TabIndex        =   410
            Top             =   495
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Pedido.frx":169776
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   13560
            TabIndex        =   397
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Aprovado por"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   69
            Left            =   8445
            TabIndex        =   328
            Top             =   300
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Aprovação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   68
            Left            =   7290
            TabIndex        =   327
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Validado por"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   67
            Left            =   5910
            TabIndex        =   326
            Top             =   300
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Validação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   66
            Left            =   4335
            TabIndex        =   325
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   65
            Left            =   2415
            TabIndex        =   324
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   63
            Left            =   10680
            TabIndex        =   323
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "N° Pedido"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   62
            Left            =   210
            TabIndex        =   322
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Emissão"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   1095
            TabIndex        =   290
            Top             =   300
            Width           =   570
         End
      End
   End
End
Attribute VB_Name = "frmCompras_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalNota                 As Double 'OK
Dim Novo_PC                   As Boolean 'OK
Public Novo_PC1                  As Boolean 'OK
Dim Novo_PC1_Custo            As Boolean 'OK
Public Novo_PC2                  As Boolean 'OK
Dim Novo_PC2_Custo            As Boolean 'OK
Dim Novo_PC3                  As Boolean 'OK
Public Compras_pedido_Prod    As Boolean
Public Compras_pedido_serv    As Boolean
Dim idpedido_compra           As String 'OK
Dim StrSql_Pedido_Necessidade As String 'OK
Dim StrSql_Pedido_Solicitacao As String 'OK
Public Sql_Pedido_Localizar   As String 'OK
Public FormulaRel_Pedido      As String 'OK
Dim TBLISTA_Compras_Pedido    As ADODB.Recordset 'OK
Dim TBLISTA_Pedido_Necessidade As ADODB.Recordset 'OK
Dim TBLISTA_Pedido_Solicitacao As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=wUC982x8R54&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=60&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

ProcCorrigeForm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeForm()
On Error GoTo tratar_erro

Cmb_empresa.Visible = False
Select Case SSTab1.Tab
    Case 0:
        Frame1(4).Visible = False
        PBLista(1).Visible = False
    Case 1:
        With Frame1(4)
            .Visible = True
            '.Top = Frame1(16).Top + Frame1(16).Height
        End With
        With PBLista(1)
            .Visible = True
            '.Top = Frame5.Top + Frame5.Height
        End With
        Cmb_empresa.Visible = True
    Case 2:
        With Frame1(4)
            .Visible = False
            '.Top = Frame1(1).Top + Frame1(1).Height
        End With
        PBLista(1).Visible = False
    Case 3:
        With Frame1(4)
            .Visible = True
            '.Top = Frame1(1).Top + Frame1(1).Height - PBLista(1).Height
        End With
        With PBLista(1)
            .Visible = True
            '.Top = Frame1(4).Top + Frame1(4).Height
        End With
    Case 4:
        With Frame1(4)
            .Visible = True
            '.Top = Frame1(1).Top + Frame1(1).Height - PBLista(1).Height
        End With
        With PBLista(1)
            .Visible = True
            '.Top = Frame1(4).Top + Frame1(4).Height
        End With
    Case 5:
        With Frame1(4)
            .Visible = False
            '.Top = Frame1(11).Top + Frame1(11).Height
        End With
        PBLista(1).Visible = False
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSalvarQtd_Click()
On Error GoTo tratar_erro

If txtstatus_item.Text = "COMPRADO" Then
frmSenha.Show 1

    If txtQuantidade.Text <> "" And TXTIDLista.Text <> "" And LiberarAlteracao = True Then
        If USMsgBox("Deseja realmente alterar a quantidade desse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            StrSql = "Update compras_Pedido_Lista set Quant_Comp = '" & Replace(txtQuantidade.Text, ",", ".") & "' WHERE idLista = '" & TXTIDLista.Text & "'"
            Conexao.Execute (StrSql)
            USMsgBox "Quantidade do item alterado com sucesso!"
            ProcAtualizalista
            LiberarAlteracao = False
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnStatusPedido_Click()
On Error GoTo tratar_erro

If txtPedido = "" Or (txtStatus <> "COMPRADO" And txtStatus <> "APROVADO") Then Exit Sub
frmCompras_Pedido_Status.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_desc_Click()
On Error GoTo tratar_erro

With txtDesconto
    If Chk_desc.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_valor_desc.Value = 0
        txtvalordesconto.Locked = True
        txtvalordesconto.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_desc2_Click()
On Error GoTo tratar_erro

With txtDesconto_serv
    If Chk_desc2.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_valor_desc2.Value = 0
        txtVlrDesconto_serv.Locked = True
        txtVlrDesconto_serv.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_valor_desc_Click()
On Error GoTo tratar_erro

With txtvalordesconto
    If Chk_valor_desc.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_desc.Value = 0
        txtDesconto.Locked = True
        txtDesconto.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_valor_desc2_Click()
On Error GoTo tratar_erro

With txtVlrDesconto_serv
    If Chk_valor_desc2.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_desc2.Value = 0
        txtDesconto_serv.Locked = True
        txtDesconto_serv.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAuto_Click()
On Error GoTo tratar_erro

If chkAuto.Value = 1 Then
    ProcLiberaTabsProd
    chkManual.Value = 0
    txtNomenclatura = ""
    Procliberacampos
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAuto_serv_Click()
On Error GoTo tratar_erro

If chkAuto_serv.Value = 1 Then
    ProcLiberaTabsServ
    chkManual_serv.Value = 0
    txtCodigo = ""
    ProcLiberaCampos_serv
Else
    ProcBloqueiaCampos_serv
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkFrete_IPI_Click()
On Error GoTo tratar_erro

ProcCalculaValor False
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_Click()
On Error GoTo tratar_erro

If chkManual.Value = 1 Then
    ProcLiberaTabsProd
    chkAuto.Value = 0
    Procliberacampos
    USMsgBox ("Informe o código interno do produto."), vbInformation, "CAPRIND v5.0"
    txtNomenclatura.Text = ""
    txtNomenclatura.SetFocus
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_serv_Click()
On Error GoTo tratar_erro

If chkManual_serv.Value = 1 Then
    ProcLiberaTabsServ
    chkAuto_serv.Value = 0
    ProcLiberaCampos_serv
    USMsgBox ("Informe o código interno do serviço."), vbInformation, "CAPRIND v5.0"
    txtCodigo.Text = ""
    txtCodigo.SetFocus
Else
    ProcBloqueiaCampos_serv
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPercentual_Click()
On Error GoTo tratar_erro

If chkPercentual.Value = 1 Then
    With txtPercentualCentro
        .Text = ""
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    chkValor.Value = 0
    With txtValorCentro
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    With txtPercentualCentro
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPercentual_serv_Click()
On Error GoTo tratar_erro

If chkPercentual_serv.Value = 1 Then
    With txtPercentualCentro_Serv
        .Text = ""
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    chkValor_serv.Value = 0
    With txtValorCentro_Serv
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    With txtPercentualCentro_Serv
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkRemessa_Click()
On Error GoTo tratar_erro

With Cmb_un_com
    If chkRemessa.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkValor_Click()
On Error GoTo tratar_erro

If chkValor.Value = 1 Then
    With txtValorCentro
        .Text = ""
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    chkPercentual.Value = 0
    With txtPercentualCentro
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    With txtValorCentro
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkValor_serv_Click()
On Error GoTo tratar_erro

If chkValor_serv.Value = 1 Then
    With txtValorCentro_Serv
        .Text = ""
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    chkPercentual_serv.Value = 0
    With txtPercentualCentro_Serv
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    With txtValorCentro_Serv
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_CST_ICMS_Click()
On Error GoTo tratar_erro

ProcCalculaValor True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_CST_ICMS_DblClick()
On Error GoTo tratar_erro

If Cmb_CST_ICMS <> "" Then
    If Len(Cmb_CST_ICMS) = 3 Then QtdeTrib = 2 Else QtdeTrib = 3
    Mercadoria = Left(Cmb_CST_ICMS, 1)
    Tributacao = Right(Cmb_CST_ICMS, QtdeTrib)
    Select Case Mercadoria
        Case 0: Origem = "0 - Nacional"
        Case 1: Origem = "1 - Estrangeira importação direta"
        Case 2: Origem = "2 - Estrangeira adquirida no mercado interno"
    End Select
    
    Select Case Tributacao
        Case "00": TributacaoICMS = "00 - Tributada integralmente"
        Case "10": TributacaoICMS = "10 - Tributada e com cobrança do ICMS por substituição"
        Case "101": TributacaoICMS = "101 - Tributada pelo Simples Nacional com permissão de crédito"
        Case "102": TributacaoICMS = "102 - Tributada pelo Simples Nacional sem permissão de crédito"
        Case "103": TributacaoICMS = "103 - Isenção do ICMS no Simples Nacional para faixa de receita bruta"
        Case "20": TributacaoICMS = "20 - Com redução de base de cálculo"
        Case "201": TributacaoICMS = "201 - Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição Tributária"
        Case "202": TributacaoICMS = "202 - Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária"
        Case "203": TributacaoICMS = "203 - Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária"
        Case "30": TributacaoICMS = "30 - Isenta ou não tributada e com cobrança do ICMS por substituição tributária"
        Case "300": TributacaoICMS = "300 - Imune"
        Case "40": TributacaoICMS = "40 - Isenta"
        Case "400": TributacaoICMS = "400 - Não tributada pelo Simples Nacional"
        Case "41": TributacaoICMS = "41 - Não tributada"
        Case "50": TributacaoICMS = "50 - Suspensão"
        Case "500": TributacaoICMS = "500 - ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação"
        Case "51": TributacaoICMS = "51 - Diferimento"
        Case "60": TributacaoICMS = "60 - ICMS cobrado anteriormente por substituição tributária"
        Case "70": TributacaoICMS = "70 - Com redução de base de cálculo e cobrança do ICMS por substituição tributária"
        Case "90": TributacaoICMS = "90 - Outras"
        Case "900": TributacaoICMS = "900 - Outros"
    End Select
End If
USMsgBox ("Origem da mercadoria do ICMS: " & Origem & vbCrLf & "Tributação pelo ICMS: " & TributacaoICMS), vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_carteira_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)
ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpar
If FunVerifStatusAprovadoPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then cmdstatus.Enabled = False Else cmdstatus.Enabled = True

IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

If IDempresa <> 0 Then
'Regime = 0
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Empresa where Codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
CNPJ_Empresa = TBFIltro!CNPJ
End If
TBFIltro.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_filtrar_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With listapedido
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar2
    Select Case Cmb_opcao_lista
        Case "Validação"
            .ButtonState(9) = 0
            .ButtonState(10) = 5
        Case "Aprovação"
            .ButtonState(9) = 5
            .ButtonState(10) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_transp_Click()
On Error GoTo tratar_erro

With cmbtransporte
    .Clear
    If Cmb_tipo_transp <> "" Then
        If Cmb_tipo_transp = "Empresa" Then
            NomeTabela = "Empresa"
            NomeCampo = "Empresa"
            NomeCampo1 = "Codigo"
        Else
            NomeCampo1 = "IDcliente"
            If Cmb_tipo_transp = "Cliente" Then
                NomeTabela = "Clientes"
                NomeCampo = "NomeRazao"
            Else
                NomeTabela = "Compras_fornecedores"
                NomeCampo = "Nome_Razao"
            End If
        End If
        
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select " & NomeCampo & ", " & NomeCampo1 & " FROM " & NomeTabela & " where " & NomeCampo & " is not null group by " & NomeCampo & ", " & NomeCampo1, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            Do While TBLISTA.EOF = False
                Select Case Cmb_tipo_transp
                    Case "Cliente":
                        .AddItem TBLISTA!NomeRazao
                        .ItemData(.NewIndex) = TBLISTA!IDCliente
                    Case "Fornecedor":
                        .AddItem TBLISTA!Nome_Razao
                        .ItemData(.NewIndex) = TBLISTA!IDCliente
                    Case "Empresa":
                        .AddItem TBLISTA!Empresa
                        .ItemData(.NewIndex) = TBLISTA!CODIGO
                End Select
                TBLISTA.MoveNext
            Loop
        End If
        TBLISTA.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_un_com_Click()
On Error GoTo tratar_erro

If txtNomenclatura <> "" Then
    If txtQuantidade <> "" Then
        If cmbun <> Cmb_un_com Then
            txtQuantidade_est = FunFormataCasasDecimais(4, FunConversaoFinalUn(cmbun, Cmb_un_com, txtQuantidade, txtNomenclatura, True))
        Else
            txtQuantidade_est = FunFormataCasasDecimais(4, txtQuantidade)
        End If
        
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            txtQuantidade_PC = FunCalculaQtdePC(txtNomenclatura, txtQuantidade, True, Cmb_un_com)
        Else
            txtQuantidade_PC = ""
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

txtEspecificacoes = FunBuscaDescPadraoFamilia(cmbfamilia, txtNomenclatura, txtEspecificacoes)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_serv_Click()
On Error GoTo tratar_erro

txtDescricao_serv = FunBuscaDescPadraoFamilia(cmbFamilia_serv, txtCodigo, txtDescricao_serv)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)
If cmbfiltrarpor_necess = "Família" Then
    txtTexto_necess.Visible = False
    cmbTexto_necess.Visible = True
    ProcCarregaComboFamilia cmbTexto_necess, "familia <> 'Null' and Compras = 'True'", True
Else
    txtTexto_necess.Visible = True
    cmbTexto_necess.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)
Txtprazo_sol.Value = Date
If cmbfiltrarpor_sol = "Família" Then
    txtTexto_sol.Visible = False
    cmbTexto_sol.Visible = True
    Txtprazo_sol.Visible = False
    ProcCarregaComboFamilia cmbTexto_sol, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", True
ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
        txtTexto_sol.Visible = False
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = True
    Else
        txtTexto_sol.Visible = True
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbMoeda_Click()
On Error GoTo tratar_erro

Txt_valor_moeda = ""
If cmbMoeda = "REAL" Then Txt_valor_moeda = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbun_Click()
On Error GoTo tratar_erro

With txtQuantidade
    .Locked = False
    .TabStop = True
End With
If cmbun <> "" Then
    ProcLibera_UN_Com cmbun, Cmb_un_com
    Cmb_un_com = cmbun
    ProcBloqLibQtde
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbUn_serv_Click()
On Error GoTo tratar_erro

If cmbUn_serv <> "" Then
    ProcLibera_UN_Com cmbun, Cmb_un_com
    Cmb_un_com_serv = cmbUn_serv
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqLibQtde()
On Error GoTo tratar_erro

If txtstatus_item = "RECEBIDO" Or txtstatus_item = "RECEBIDO PARCIAL" Or txtstatus_item = "CANCELADO" Then Exit Sub

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtNomenclatura = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtNomenclatura & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo1_Click()
On Error GoTo tratar_erro

If txtCodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtCodigo & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAbrir_codigo_Click()
On Error GoTo tratar_erro

If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "serviços", "localizar", True, True) = False Then Exit Sub
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "serviços", "localizar", True, True) = False Then Exit Sub
End If
If FunVerifSatus("localizar serviços", True) = False Then Exit Sub
ProcLiberaTabsServ
Sit_Nota = 2
frmCompras_ListaProduto.Show 1
If txtDescricao_serv <> "" Then txtPrazo_serv.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDPedido = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from compras_pedido order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idpedido = " & txtIDPedido)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtIDPedido.Text = TBLISTA!IDpedido
        Set TBCompras_Pedido = CreateObject("adodb.recordset")
        TBCompras_Pedido.Open "Select * from compras_pedido where idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        ProcLimparComercial
        TXTIDLista = 0
        ProcLimpaCamposItem True
        ProcLimpaCamposCusto
        SSTab2.Tab = 0
        txtIDLista_serv = 0
        ProcLimpaCamposServ True
        ProcLimpaCamposCustoServ
        SSTab3.Tab = 0
        ProcPuxaDados
        ProcAbreComercial
        ProcCarregaEscopoForn
        ProcAtualizalista
        ProcAtualizalistaServ
    Else
        USMsgBox ("Fim dos cadastros de pedido de compra."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_PC1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviarEmail()
On Error GoTo tratar_erro

Acao = "enviar o e-mail"
If txtIDPedido = 0 Then
    NomeCampo = "o pedido de compra"
    ProcVerificaAcao
    Exit Sub
End If

If txtStatus <> "APROVADO" And txtStatus <> "COMPRADO" Then
    USMsgBox ("Não é permitido enviar o e-mail com pedido de compra no status de " & txtStatus.Text & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Compras_Pedido = True
Custos_justificativa = False
Vendas_Proposta = False

 Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select email from Empresa_email where Usuario_caprind ='" & pubUsuario & "' AND Aplicacao = 'C'", Conexao, adOpenKeyset, adLockOptimistic
     If TBFIltro.EOF = False Then
      If IsNull(TBFIltro!Email) Then
       USMsgBox ("Atenção usuário " & pubUsuario & ", não foi encontrado um servidor de email valido no seu cadastro!"), vbCritical, "CAPRIND v5.0"
       Exit Sub
      End If
     FrmEnviarEmail.txtDe.Text = TBFIltro!Email
     End If
    TBFIltro.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Caminho from Empresa_armazenamento_PDF where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Relatorio = 'Pedido de compra'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    NomeRel = "Compras_pedido.rpt"
    If Len(TBAbrir!caminho) = 3 Then caminho = TBAbrir!caminho Else caminho = TBAbrir!caminho & "\"
    Nome_anexo = Replace(txtPedido, "/", "-") & ".pdf"
    ProcGerarPDF caminho & Nome_anexo, "{compras_pedido.IDpedido} = " & txtIDPedido, ""
    
    FrmEnviarEmail.Txt_anexo = caminho & Nome_anexo
    FrmEnviarEmail.lblanexo.Caption = "Pedido de compra n°" & Nome_anexo
Else
    USMsgBox ("Atenção usuário " & pubUsuario & ", não foi encontrado um local para armazenamento de documentos PDF válido!"), vbCritical, "CAPRIND v5.0"
    Exit Sub
End If
TBAbrir.Close
FrmEnviarEmail.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362P" Then frmCompras_Pedido_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmCompras_Pedido_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza dados dos pedidos
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                TBGravar.MoveLast
                PBLista(1).Min = 0
                PBLista(1).Max = TBGravar.RecordCount
                PBLista(1).Value = 1
                Contador = 0
                TBGravar.MoveFirst
                Do While TBGravar.EOF = False
                    If IsNull(TBGravar!Ordem) = True Or TBGravar!Ordem = "" Then TBGravar!Ordem = 0
                    If IsNull(TBGravar!IDpedido) = True Or TBGravar!IDpedido = "" Then TBGravar!IDpedido = 0
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where desenho = '" & TBGravar!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        If TBProduto!Tipo = "P" Or TBProduto!Tipo = "I" Then TBGravar!Tipo = "P" Else TBGravar!Tipo = "S"
                        TBGravar!Codproduto = TBProduto!Codproduto
                    Else
                        If IsNull(TBGravar!Tipo) = True Or TBGravar!Tipo = "" Then TBGravar!Tipo = "P"
                    End If
                    TBProduto.Close
                    TBGravar.Update
                    TBGravar.MoveNext
                    Contador = Contador + 1
                    PBLista(1).Value = Contador
                Loop
            End If
            TBGravar.Close
        End If
            
        If .Chk2.Value = 1 Then
            'Atualiza status dos pedidos
            'Altera status do item
            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido_lista where idpedido <> 0 and (Status_Item = 'RECEBIDO' or Status_Item = 'N_RECEBIDO' or Status_Item = 'PARCIAL') order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                TBCompras_Pedido.MoveLast
                PBLista(1).Min = 0
                PBLista(1).Max = TBCompras_Pedido.RecordCount
                PBLista(1).Value = 1
                Contador = 0
                TBCompras_Pedido.MoveFirst
                Do While TBCompras_Pedido.EOF = False
                    If IsNull(TBCompras_Pedido!IDpedido) = False And TBCompras_Pedido!IDpedido <> "" Then
                        Qtd = 0
                        Set TBEstoque = CreateObject("adodb.recordset")
                        TBEstoque.Open "Select Sum(Recebido) as Qtd from Estoque_controle_recebimento where Idlista = " & TBCompras_Pedido!IDlista & " and idpedido = " & TBCompras_Pedido!IDpedido & " and Desenho = '" & TBCompras_Pedido!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBEstoque.EOF = False Then
                            Qtd = IIf(IsNull(TBEstoque!Qtd), 0, TBEstoque!Qtd)
                        End If
                        TBEstoque.Close
                        If TBCompras_Pedido!Status_Item <> "CANCELADO" Then
                            If Qtd = 0 Then
                                TBCompras_Pedido!Status_Item = "N_RECEBIDO"
                            ElseIf Qtd < TBCompras_Pedido!Quant_Comp Then
                                    TBCompras_Pedido!Status_Item = "PARCIAL"
                                ElseIf Qtd >= TBCompras_Pedido!Quant_Comp Then
                                        TBCompras_Pedido!Status_Item = "RECEBIDO"
                            End If
                        End If
                        TBCompras_Pedido.Update
                    End If
                    TBCompras_Pedido.MoveNext
                    Contador = Contador + 1
                    PBLista(1).Value = Contador
                Loop
            End If
            TBCompras_Pedido.Close
                    
            'Altera status do pedido
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from Compras_pedido order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                TBPedido.MoveLast
                PBLista(1).Min = 0
                PBLista(1).Max = TBPedido.RecordCount
                PBLista(1).Value = 1
                Contador = 0
                TBPedido.MoveFirst
                Do While TBPedido.EOF = False
                    Aberto = False
                    Recebido = False
                    Parcial = False
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select * from compras_pedido_lista where idpedido = " & TBPedido!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        Do While TBItem.EOF = False
                            If TBItem!Status_Item = "RECEBIDO" Then Recebido = True
                            If TBItem!Status_Item = "PARCIAL" Then Parcial = True
                            If TBItem!Status_Item = "N_RECEBIDO" Then Aberto = True
                            TBItem.MoveNext
                        Loop
                        If Aberto = False And Parcial = False And Recebido = True Then TBPedido!Status_pedido = "ENCERRADO"
                        If Aberto = True And Parcial = False And Recebido = False Then TBPedido!Status_pedido = "ABERTO"
                        If Parcial = True Or Aberto = True And Recebido = True Then TBPedido!Status_pedido = "PARCIAL"
                    Else
                        TBPedido!Status_pedido = "ABERTO"
                    End If
                    TBItem.Close
                    TBPedido.Update
                    TBPedido.MoveNext
                    Contador = Contador + 1
                    PBLista(1).Value = Contador
                Loop
            End If
            TBPedido.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Centro de custo
            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido_lista where Centro <> 'Null' order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                TBCompras_Pedido.MoveLast
                PBLista(1).Min = 0
                PBLista(1).Max = TBCompras_Pedido.RecordCount
                PBLista(1).Value = 1
                Contador = 0
                TBCompras_Pedido.MoveFirst
                Do While TBCompras_Pedido.EOF = False
                    If TBCompras_Pedido!centro <> "" Then
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Compras_pedido_lista_custo where IDLista = " & TBCompras_Pedido!IDlista, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = True Then TBGravar.AddNew
                        TBGravar!IDpedido = TBCompras_Pedido!IDpedido
                        TBGravar!IDlista = TBCompras_Pedido!IDlista
                        TBGravar!CentroCusto = TBCompras_Pedido!centro
                        TBGravar!Responsavel = pubUsuario
                        TBGravar!Data = Date
                        TBGravar!valor = TBCompras_Pedido!preco_total
                        TBGravar!Percentual = 100
                        TBGravar.Update
                        TBGravar.Close
                    End If
                    TBCompras_Pedido.MoveNext
                    Contador = Contador + 1
                    PBLista(1).Value = Contador
                Loop
            End If
            TBCompras_Pedido.Close
        End If
        
        If .Chk4.Value = 1 Then
            'Numero do endereço do fornecedor nos pedidos
            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                TBCompras_Pedido.MoveLast
                PBLista(1).Min = 0
                PBLista(1).Max = TBCompras_Pedido.RecordCount
                PBLista(1).Value = 1
                Contador = 0
                TBCompras_Pedido.MoveFirst
                Do While TBCompras_Pedido.EOF = False
                    If IsNull(TBCompras_Pedido!Numero) = True Or TBCompras_Pedido!Numero = "" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Compras_fornecedores where IDCliente = " & IIf(IsNull(TBCompras_Pedido!IDFornecedor), 0, TBCompras_Pedido!IDFornecedor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            TBCompras_Pedido!Numero = IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero)
                            TBCompras_Pedido.Update
                        End If
                    End If
                    TBCompras_Pedido.MoveNext
                    Contador = Contador + 1
                    PBLista(1).Value = Contador
                Loop
            End If
            TBCompras_Pedido.Close
        End If
        
        If .Chk5.Value = 1 Then
            'Fornecedor nos produtos/serviços
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select Compras_pedido.idfornecedor, Compras_pedido_lista.* from Compras_pedido INNER JOIN Compras_pedido_lista ON Compras_pedido.IDpedido = Compras_pedido_lista.IDpedido order by Compras_pedido.IDpedido", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                TBPedido.MoveLast
                PBLista(1).Min = 0
                PBLista(1).Max = TBPedido.RecordCount
                PBLista(1).Value = 1
                Contador = 0
                TBPedido.MoveFirst
                Do While TBPedido.EOF = False
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select codproduto from Projproduto where desenho = '" & TBPedido!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBPedido!Codproduto = TBFI!Codproduto
                        TBPedido.Update
                        Set TBOrdem = CreateObject("adodb.recordset")
                        TBOrdem.Open "Select Projproduto_fornecedor.* from Projproduto_fornecedor INNER JOIN Projproduto on Projproduto_fornecedor.codproduto = Projproduto.codproduto where Projproduto.desenho = '" & TBPedido!Desenho & "' and Projproduto_fornecedor.Idfornecedor = " & TBPedido!IDFornecedor, Conexao, adOpenKeyset, adLockOptimistic
                        If TBOrdem.EOF = True Then
                            TBOrdem.AddNew
                            TBOrdem!Codproduto = TBFI!Codproduto
                            TBOrdem!IDFornecedor = TBPedido!IDFornecedor
                        End If
                        TBOrdem!PCusto = TBPedido!preco_unitario
                        TBOrdem.Update
                        TBOrdem.Close
                    End If
                    TBFI.Close
                    TBPedido.MoveNext
                    Contador = Contador + 1
                    PBLista(1).Value = Contador
                Loop
            End If
            TBPedido.Close
        End If
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Compras/Pedido"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCalculadora()
On Error GoTo tratar_erro

If txtNomenclatura = "" Then
    USMsgBox ("Informe o código interno antes de abrir a calculadora para cálculo de peso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtstatus_item <> "AGUARDANDO APROVAÇÃO" And txtstatus_item.Text <> "COMPRADO" Then
    USMsgBox ("Não é permitido abrir a calculadora para cálculo de peso, pois este produto está " & txtstatus_item & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = False
    Compras_Cotacao = False
    Compras_Pedido = True
    Estoque_recebimento = False
    Vendas_Proposta = False
    Vendas_PI = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterarStatusItem()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If TXTIDLista = 0 Then
    Acao = "alterar o status"
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
If txtstatus_item <> "AGUARDANDO APROVAÇÃO" And txtstatus_item.Text <> "APROVADO" And txtstatus_item.Text <> "COMPRADO" And txtstatus_item.Text <> "CANCELADO" Then
    USMsgBox ("Não é permitido alterar o status deste produto, pois o mesmo está " & txtstatus_item & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If (txtstatus_item = "COMPRADO" Or txtstatus_item = "APROVADO") And txtResponsavel_aprovacao <> pubUsuario Then
    USMsgBox ("Somente o usuário que aprovou o pedido pode alterar o status deste produto."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then If FunVerifValidacaoRegistro("alterar status", txtDtValidacao, "pedido de compra", "do produto", True) = False Then Exit Sub
Else
    If FunVerifValidacaoRegistro("alterar status", txtDtValidacao, "pedido de compra", "do produto", True) = False Then Exit Sub
End If

If txtstatus_item = "CANCELADO" Then GoTo 1
If USMsgBox("Deseja realmente alterar o status deste produto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
1:
    If FunVerifCancelamento("CPL.IDlista = " & TXTIDLista, True, False) = False Then Exit Sub
    IDlista = txtIDPedido
    Compras_Pedido = True
    Compras_pedido_Prod = True
    Compras_pedido_serv = False
    Vendas_Proposta = False
    Vendas_PI = False
    Plano_centro_de_custo = False
    frmCompras_pedido_cancelar.Show 1
    TXTIDLista = 0
    ProcLimpaCamposItem True
    Frame1(12).Enabled = False
    ProcAtualizalista
    Novo_PC1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterarStatusServico()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If txtIDLista_serv = 0 Then
    Acao = "alterar o status"
    NomeCampo = "o serviço"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus_serv <> "AGUARDANDO APROVAÇÃO" And txtStatus_serv.Text <> "APROVADO" And txtStatus_serv.Text <> "COMPRADO" And txtStatus_serv.Text <> "CANCELADO" Then
    USMsgBox ("Não é permitido alterar o status deste serviço, pois o mesmo está " & txtStatus_serv & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If (txtStatus_serv = "COMPRADO" Or txtStatus_serv.Text <> "APROVADO") And txtResponsavel_aprovacao <> pubUsuario Then
    USMsgBox ("Somente o usuário que aprovou o pedido pode alterar o status deste produto."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then If FunVerifValidacaoRegistro("alterar status", txtDtValidacao, "pedido de compra", "do serviço", True) = False Then Exit Sub
Else
    If FunVerifValidacaoRegistro("alterar status", txtDtValidacao, "pedido de compra", "do serviço", True) = False Then Exit Sub
End If

If txtStatus_serv = "CANCELADO" Then GoTo 1
If USMsgBox("Deseja realmente alterar o status deste serviço?", vbYesNo, "CAPRIND v5.0") = vbYes Then
1:
    If FunVerifCancelamento("CPL.IDlista = " & txtIDLista_serv, False, True) = False Then Exit Sub
    IDlista = txtIDPedido
    Compras_Pedido = True
    Compras_pedido_Prod = False
    Compras_pedido_serv = True
    Vendas_Proposta = False
    Vendas_PI = False
    Plano_centro_de_custo = False
    frmCompras_pedido_cancelar.Show 1
    txtIDLista_serv = 0
    ProcLimpaCamposServ True
    Frame1(7).Enabled = False
    ProcAtualizalistaServ
    Novo_PC2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdCF_Click()
On Error GoTo tratar_erro

Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Clientes = False
Compras_Pedido = True
Familia_NCM = False
ClassFiscal = False
frmProj_Classificacao_Fiscal.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCFOP_prod_Click()
On Error GoTo tratar_erro

Clientes = False
Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Compras_Pedido = True
Sit_REG = 1
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCFOP_serv_Click()
On Error GoTo tratar_erro

Clientes = False
Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Compras_Pedido = True
Sit_REG = 2
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCondicoes_Click()
On Error GoTo tratar_erro

Aplic = 1
Compras_Cotacao = False
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcontatos_Click()
On Error GoTo tratar_erro

If txtIDfornecedor <> "" And txtIDfornecedor <> "0" Then
    Compras_Cotacao = False
    Compras_Pedido = True
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    frmCompras_Pedido_contatos.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtPedido = "" Then
    USMsgBox ("Informe o pedido interno antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente copiar o pedido de compra " & txtPedido.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Revisar = False
    ProcCopiarPedido
    '==================================
    Modulo = "Compras/Pedido de compra"
    Evento = "Novo"
    ID_documento = txtIDPedido
    Documento = "Nº pedido: " & txtPedido
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from compras_pedido where idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        ProcLimpar
        ProcLimpaCamposItem True
        ProcLimpaCamposServ True
        ProcPuxaDados
    End If
    TBCompras_Pedido.Close
    
    ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(lblPaginas(3).Caption, Len(lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(3).Caption, Len(lblPaginas(3).Caption) - 5))))
    Frame1(16).Enabled = True
    Frame1(4).Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExportarExcel()
On Error GoTo tratar_erro
Dim exclApp As Object
Dim exclBook As Object
Dim excSheet As Object

'If Alterar = False Then
'    usMsgbox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
Acao = "exportar"
If txtIDPedido = 0 Then
    NomeCampo = "o pedido de compra"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus = "AGUARDANDO APROVAÇÃO" Then
    USMsgBox ("Não é permitido exportar, pois o pedido de compra ainda não foi aprovado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente exportar este pedido de compra " & txtPedido.Text & " para um arquivo em excel?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    caminho = Localrel & "\Arquivos exportados\Pedido de compra"
    If GerArqPastas.FolderExists(caminho) = False Then
        USMsgBox ("Não é permitido exportar, pois não foi encontrado o caminho " & caminho & ", onde será armazenado o aquivo."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    'Carregar o Excel:
    Set exclApp = CreateObject("Excel.Application")
    'Crie um WorkBook:
    Set exclBook = exclApp.Workbooks.Add

    'Defina a planilha ativa p/ facilitar o trabalho:
    Set exclSheet = exclApp.ActiveWorkbook.ActiveSheet

    With exclSheet
        'Definir o conteúdo (Linha, Coluna):
        
        .Cells(1, 1).Value = "INÍCIO"
        
        'EMPRESA
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Cells(2, 1).Value = TBAbrir!Razao
            .Cells(2, 2).Value = TBAbrir!CNPJ
            .Cells(2, 3).Value = TBAbrir!IE
            .Cells(2, 4).Value = TBAbrir!IM
            .Cells(2, 5).Value = TBAbrir!Tipo_endereco
            .Cells(2, 6).Value = TBAbrir!Endereco
            .Cells(2, 7).Value = TBAbrir!Numero
            .Cells(2, 8).Value = TBAbrir!complemento
            .Cells(2, 9).Value = TBAbrir!Tipo_bairro
            .Cells(2, 10).Value = TBAbrir!Bairro
            .Cells(2, 11).Value = TBAbrir!Cidade
            .Cells(2, 12).Value = TBAbrir!CEP
            .Cells(2, 13).Value = TBAbrir!UF
            .Cells(2, 14).Value = TBAbrir!Pais
            .Cells(2, 15).Value = TBAbrir!Codigo_pais
            .Cells(2, 16).Value = TBAbrir!telefone
            .Cells(2, 17).Value = TBAbrir!Fax
            .Cells(2, 18).Value = TBAbrir!Email
            .Cells(2, 19).Value = TBAbrir!Site
            If TBAbrir!Simples = True Then
                .Cells(2, 29).Value = "Simples"
            ElseIf TBAbrir!Real = True Then
                    .Cells(2, 20).Value = "Real"
                Else
                    .Cells(2, 20).Value = "Presumido"
            End If
        End If
        TBAbrir.Close
        
        'DADOS COMERCIAIS
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_comercial where IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Cells(3, 1).Value = TBAbrir!condicoes
            .Cells(3, 2).Value = TBAbrir!Embalagem
            .Cells(3, 3).Value = TBAbrir!Observacoes
            .Cells(3, 4).Value = TBAbrir!Prazo
            .Cells(3, 5).Value = TBAbrir!localentrega
            .Cells(3, 6).Value = TBAbrir!Escopo
            .Cells(3, 7).Value = TBAbrir!Moeda
            .Cells(3, 8).Value = TBAbrir!Valor_moeda
        End If
        TBAbrir.Close
        
        'RESPONSÁVEL
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Responsavel from Compras_pedido where IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Cells(3, 9).Value = TBAbrir!Responsavel
        End If
        TBAbrir.Close
        
        'LISTA DE PRODUTOS
        Contador = 4
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_pedido_lista where IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                .Cells(Contador, 1).Value = TBAbrir!Desenho
                .Cells(Contador, 2).Value = TBAbrir!Descricao
                .Cells(Contador, 3).Value = TBAbrir!Quant_Comp
                .Cells(Contador, 4).Value = TBAbrir!preco_unitario
                .Cells(Contador, 5).Value = TBAbrir!IPI
                .Cells(Contador, 6).Value = TBAbrir!preco_total
                .Cells(Contador, 7).Value = TBAbrir!Un
                .Cells(Contador, 8).Value = TBAbrir!vlrICMS
                .Cells(Contador, 9).Value = TBAbrir!VlrIPI
                .Cells(Contador, 10).Value = TBAbrir!ICMS
                .Cells(Contador, 11).Value = TBAbrir!Familia
                .Cells(Contador, 12).Value = TBAbrir!Obs
                .Cells(Contador, 13).Value = TBAbrir!Prazo
                .Cells(Contador, 14).Value = TBAbrir!Desconto
                .Cells(Contador, 15).Value = TBAbrir!ValorDesconto
                .Cells(Contador, 16).Value = TBAbrir!preco_unitario_desconto
                .Cells(Contador, 17).Value = TBAbrir!Remessa
                .Cells(Contador, 18).Value = TBAbrir!Tipo
                .Cells(Contador, 19).Value = TBAbrir!ISSQN
                .Cells(Contador, 20).Value = TBAbrir!VlrISSQN
                .Cells(Contador, 21).Value = TBAbrir!N_referencia
                .Cells(Contador, 22).Value = TBAbrir!Descricao_comercial
                .Cells(Contador, 23).Value = TBAbrir!Unidade_com
                
                'CFOP
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBAbrir!ID_CFOP), 0, TBAbrir!ID_CFOP), Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    .Cells(Contador, 24).Value = TBFI!ID_CFOP
                    .Cells(Contador, 25).Value = TBFI!Txt_descricao
                    .Cells(Contador, 26).Value = TBFI!Txt_ICMS
                    .Cells(Contador, 27).Value = TBFI!txt_IPI
                    .Cells(Contador, 28).Value = TBFI!txt_Somar
                    .Cells(Contador, 29).Value = TBFI!Vendas
                    .Cells(Contador, 30).Value = TBFI!Retem
                    .Cells(Contador, 31).Value = TBFI!Suframa
                    .Cells(Contador, 32).Value = TBFI!MaoObra
                    .Cells(Contador, 33).Value = TBFI!Demonstracao
                    .Cells(Contador, 34).Value = TBFI!Soma_retorno_totalnf
                    .Cells(Contador, 35).Value = TBFI!TemPIS
                    .Cells(Contador, 36).Value = TBFI!TemCOFINS
                    .Cells(Contador, 37).Value = TBFI!De
                    .Cells(Contador, 38).Value = TBFI!FE
                    .Cells(Contador, 39).Value = TBFI!MPA
                    .Cells(Contador, 40).Value = TBFI!TemReducaoBC
                    .Cells(Contador, 41).Value = TBFI!Remessa
                    .Cells(Contador, 42).Value = TBFI!retorno
                    .Cells(Contador, 43).Value = TBFI!Somar_IPI_BC_ICMSST
                    .Cells(Contador, 44).Value = TBFI!Devolucao
                End If
                TBFI.Close
                
                'NCM
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from tbl_ClassificacaoFiscal where IdClass = " & IIf(IsNull(TBAbrir!ID_CF), 0, TBAbrir!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    .Cells(Contador, 45).Value = TBFI!IDIntClasse
                    .Cells(Contador, 46).Value = TBFI!Txt_grupo
                    .Cells(Contador, 47).Value = TBFI!txt_Class
                    .Cells(Contador, 48).Value = TBFI!dbl_ICMS_de
                    .Cells(Contador, 49).Value = TBFI!dbl_ICMS_ss
                    .Cells(Contador, 50).Value = TBFI!dbl_ICMS_nn
                    .Cells(Contador, 51).Value = TBFI!dbl_ICMS_co
                    .Cells(Contador, 52).Value = TBFI!dbl_IPI
                    .Cells(Contador, 53).Value = TBFI!basereduzida
                    .Cells(Contador, 54).Value = TBFI!CTDE
                    .Cells(Contador, 55).Value = TBFI!CTNN
                    .Cells(Contador, 56).Value = TBFI!CTCO
                    .Cells(Contador, 57).Value = TBFI!CTSS
                    .Cells(Contador, 58).Value = TBFI!Retem_PIS_Cofins
                    .Cells(Contador, 59).Value = TBFI!PIS
                    .Cells(Contador, 60).Value = TBFI!Cofins
                    .Cells(Contador, 61).Value = TBFI!PIS_destaca
                    .Cells(Contador, 62).Value = TBFI!Cofins_destaca
                    .Cells(Contador, 63).Value = TBFI!dbl_ICMS_ex
                    .Cells(Contador, 64).Value = TBFI!CTEX
                    .Cells(Contador, 65).Value = TBFI!Desoneracao
                    .Cells(Contador, 66).Value = TBFI!Aliq_nacional
                    .Cells(Contador, 67).Value = TBFI!Aliq_importacao
                End If
                
                .Cells(Contador, 68).Value = TBAbrir!CST
                .Cells(Contador, 69).Value = TBAbrir!Valor_ICMS_ST
                .Cells(Contador, 70).Value = TBAbrir!BC_ICMS_ST
                .Cells(Contador, 71).Value = TBAbrir!BC_ICMS
                .Cells(Contador, 72).Value = TBAbrir!Frete
                .Cells(Contador, 73).Value = TBAbrir!Seguro
                .Cells(Contador, 74).Value = TBAbrir!Acessorias
                .Cells(Contador, 75).Value = TBAbrir!Frete_IPI
                
                'ESTRUTURA
                If TBAbrir!Tipo = "P" And TBAbrir!Remessa = False Then
                    If TBAbrir!Desenho = "00002-FC" Then
                        USMsgBox "Aqui"
                    End If
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from Projconjunto where Codproduto = " & TBAbrir!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Do While TBFI.EOF = False
                            'Verifica se existe o material no pedido como remessa
                            Set TBFIltro = CreateObject("adodb.recordset")
                            TBFIltro.Open "Select IDlista from Compras_pedido_lista where IDpedido = " & TBAbrir!IDpedido & " and Desenho = '" & TBFI!Desenho & "' and Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFIltro.EOF = False Then
                                .Cells(Contador, 76).Value = TBFI!Desenho
                                .Cells(Contador, 77).Value = TBFI!Descricao
                                .Cells(Contador, 78).Value = TBFI!PesoMetro
                                .Cells(Contador, 79).Value = TBFI!PesoTotal
                                .Cells(Contador, 80).Value = TBFI!quantidade
                                .Cells(Contador, 81).Value = TBFI!Peso
                                .Cells(Contador, 82).Value = TBFI!Unidade
                                .Cells(Contador, 83).Value = TBFI!Dimensoes
                                .Cells(Contador, 84).Value = TBFI!Un_Kg
                                .Cells(Contador, 85).Value = TBFI!Posicao
                            End If
                            TBFIltro.Close
                            TBFI.MoveNext
                        Loop
                    End If
                    TBFI.Close
                End If
                
                Contador = Contador + 1
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        
        .Cells(Contador, 1).Value = "FIM"
        .SaveAs Localrel & "\Arquivos exportados\Pedido de compra\" & Replace(txtPedido, "/", "-") & ".xls"
    End With
    
    exclApp.Workbooks.Close
    'Limpe as variáveis de Objeto:
    Set exclSheet = Nothing
    Set exclBook = Nothing
    Set exclApp = Nothing
    
    USMsgBox ("Pedido de compra exportado com sucesso para pasta " & caminho & "."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Compras/Pedido de compra"
    Evento = "Exportar para arquivo em excel"
    ID_documento = txtIDPedido
    Documento = "Nº pedido: " & txtPedido
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDados_pagto_Click()
On Error GoTo tratar_erro

frmCompras_Pedido_banco.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDadosComerciaisFornecedor_Click()
On Error GoTo tratar_erro


Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM Compras_fornecedores_DadosComerciais WHERE IDfornecedor = " & IIf(txtIDfornecedor = "", 0, txtIDfornecedor) & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
'    txtcalculos.Text = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
'    txt.Text = IIf(IsNull(TBCotacao!impostos), "", TBCotacao!impostos)
    cmbpagamento.Text = IIf(IsNull(TBCotacao!condicoes), "", TBCotacao!condicoes)
'    txtgarantia.Text = IIf(IsNull(TBCotacao!garantia), "", TBCotacao!garantia)
'    txtReajuste.Text = IIf(IsNull(TBCotacao!reajuste), "", TBCotacao!reajuste)
    txttransporte.Text = IIf(IsNull(TBCotacao!transporte), "", TBCotacao!transporte)
'    txtValidade.Text = IIf(IsNull(TBCotacao!validade), "", TBCotacao!validade)
'    txtID_cfop = IIf(IsNull(TBCotacao!IDCFOP), "", TBCotacao!IDCFOP)
'    txtCFOP = IIf(IsNull(TBCotacao!CFOP), "", TBCotacao!CFOP)
'    txtoperacao = IIf(IsNull(TBCotacao!descricaoCFOP), "", TBCotacao!descricaoCFOP)
Else
USMsgBox "Não existe dados comerciais cadastrados para esse fornecedor", vbInformation, "CAPRIND v5.0", "CAPRIND"
End If
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEmbalagem_Click()
On Error GoTo tratar_erro

Aplic = 3
Compras_Cotacao = False
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEntrega_Click()
On Error GoTo tratar_erro

Aplic = 2
Compras_Cotacao = False
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If txtIDPedido = 0 Then
    USMsgBox ("Informe o pedido de compra antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus.Text <> "AGUARDANDO APROVAÇÃO" And txtStatus <> "CANCELADO" Then
    USMsgBox ("Não é permitido alterar o status do pedido de compra, pois o mesmo está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus = "CANCELADO" Then GoTo 1
If USMsgBox("Deseja realmente alterar o status do pedido de compra nº " & txtPedido.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
1:
    If FunVerifCancelamento("CP.IDpedido = " & txtIDPedido, False, False) = False Then Exit Sub
    
    IDlista = txtIDPedido
    Compras_Pedido = True
    Compras_pedido_Prod = False
    Compras_pedido_serv = False
    Vendas_Proposta = False
    Vendas_PI = False
    Plano_centro_de_custo = False
    frmCompras_pedido_cancelar.Show 1
    ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(lblPaginas(3).Caption, Len(lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(3).Caption, Len(lblPaginas(3).Caption) - 5))))
    Novo_PC = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifCancelamento(TextoFiltro As String, Prod As Boolean, Serv As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifCancelamento = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CPLE.ID from (Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CPL.IDpedido = CP.IDpedido) INNER JOIN Compras_pedido_lista_empenhos CPLE ON CPLE.IDlista = CPL.IDlista where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Prod = True Or Serv = True Then
        MsgTexto = IIf(Prod = True, "produto", "serviço")
        MsgTexto1 = "pois o mesmo está empenhado"
    Else
        MsgTexto = "pedido"
        MsgTexto1 = "pois exitem(m) produto(s)/serviço(s) empenhado(s)"
    End If
    USMsgBox ("Não é permitido alterar o status deste " & MsgTexto & ", " & MsgTexto1 & "."), vbExclamation, "CAPRIND v5.0"
    FunVerifCancelamento = False
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcExcluirTab_Servico()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab3.Tab
    Case 0: ProcExcluirServ
    Case 1: ProcExcluirServ_custo
    Case 2: ProcExcluirEmpenhoServ
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

ProcLimpaCamposItem False
ProcCarregaDadosItem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosItem()
On Error GoTo tratar_erro

chkRemessa.Enabled = True
If txtNomenclatura <> "" Then
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select * from projproduto where desenho = '" & txtNomenclatura.Text & "' and Tipo = 'P' and DtValidacao IS NOT NULL and (Compras = 'True' or Producao = 'True')", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        txtcodproduto = 0
        cmbun.ListIndex = -1
        Cmb_un_com.ListIndex = -1
        cmbfamilia.ListIndex = -1
        txtvalorunitario = "0,00000"
        txtEspecificacoes.Text = ""
        txtDescricao_comercial.Text = ""
        cmbReferencia.Clear
        
        If TBCompras_Lista!Bloqueado = False Then
            txtcodproduto = TBCompras_Lista!Codproduto
            txtNomenclatura = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
            txtEspecificacoes.Text = IIf(IsNull(TBCompras_Lista!Descricao), "", (TBCompras_Lista!Descricao))
            txtDescricao_comercial.Text = IIf(IsNull(TBCompras_Lista!descricaotecnica), "", (TBCompras_Lista!descricaotecnica))
            If IsNull(TBCompras_Lista!Unidade) = False And TBCompras_Lista!Unidade <> "" Then cmbun.Text = TBCompras_Lista!Unidade
            'If IsNull(TBCompras_Lista!Unidade) = False And TBCompras_Lista!Unidade <> "" Then Cmb_un_com.Text = TBCompras_Lista!Unidade
            If IsNull(TBCompras_Lista!Unidade_com) = False And TBCompras_Lista!Unidade_com <> "" Then Cmb_un_com.Text = TBCompras_Lista!Unidade_com
            If IsNull(TBCompras_Lista!Classe) = False And TBCompras_Lista!Classe <> "" Then cmbfamilia.Text = TBCompras_Lista!Classe
2:
            valor = IIf(Txt_valor_moeda = "", 1, Txt_valor_moeda)
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Projproduto_fornecedor where Codproduto = " & TBCompras_Lista!Codproduto & " and idfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                txtvalorunitario = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto / valor, "###,##0.0000000000"))
            Else
                txtvalorunitario = IIf(IsNull(TBCompras_Lista!PCusto), "", (Format(TBCompras_Lista!PCusto / valor, "###,##0.0000000000")))
            End If
            TBFI.Close
            
            If TBCompras_Lista!Estoque = True Then
                With cmbun
                    .Locked = True
                    .TabStop = False
                End With
    '            With Cmb_un_com
    '                .Locked = True
    '                .TabStop = False
    '            End With
            Else
                With cmbun
                    .Locked = False
                    .TabStop = True
                End With
    '            With Cmb_un_com
    '                .Locked = False
    '                .TabStop = True
    '            End With
            End If
            
            With chkRemessa
                If TBCompras_Lista!Compras = False Then
                    .Value = 1
                    .Enabled = False
                Else
                    .Value = 0
                    .Enabled = True
                End If
            End With
            
            If IsNull(TBCompras_Lista!ID_CF) = False Then Txt_ID_CF = TBCompras_Lista!ID_CF
            ProcCarregaDadosCFOPProdServ IIf(IsNull(TBCompras_Lista!ID_CFOP), 0, TBCompras_Lista!ID_CFOP), True
            
            If Txt_ID_CF <> "" Then
                ProcValorImposto txtPedido, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), 0
                ProcControleImposto IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), 0
                If TemIPI = "SIM" Then txtIPI = IntIPI Else txtIPI.Text = 0
                If TemICMS = "SIM" Then txtICMS = IntICMS Else txtICMS = 0
            End If
            
            ProcCarregaComboCodRef cmbReferencia, "P.desenho = '" & txtNomenclatura & "'", txtIDfornecedor, "F", True, True
            If TBCompras_Lista!Valor_bloqueado = True Then txtvalorunitario.Locked = True Else txtvalorunitario.Locked = False
            
            Txt_vlr_unit_ultima_compra_prod = FunVerifVlrUnitUltCompra(txtNomenclatura, 0)
        Else
            USMsgBox ("Não é permitido utilizar este produto, pois o mesmo está bloqueado."), vbExclamation, "CAPRIND v5.0"
        End If
        ProcBloqueiaCamposProdComCadastrado
    Else
        TXTIDLista = 0
        ProcLimpaCamposItem False
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposProdSemCadastrado
    End If
    TBCompras_Lista.Close
Else
    TXTIDLista = 0
    ProcLimpaCamposItem True
    If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposProdSemCadastrado
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a unidade ou familia desse registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposProdComCadastrado()
On Error GoTo tratar_erro

With txtEspecificacoes
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercial
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With
'With Cmb_un_com
'    .Locked = True
'    .TabStop = False
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposProdSemCadastrado()
On Error GoTo tratar_erro

With txtEspecificacoes
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercial
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
With cmbun
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .Locked = False
    .TabStop = True
End With
With txtvalorunitario
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposServComCadastrado()
On Error GoTo tratar_erro

With txtDescricao_serv
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercialServ
    .Locked = True
    .TabStop = False
End With
With cmbFamilia_serv
    .Locked = True
    .TabStop = False
End With
With cmbUn_serv
    .Locked = True
    .TabStop = False
End With
'With Cmb_un_com_serv
'    .Locked = True
'    .TabStop = False
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposServSemCadastrado()
On Error GoTo tratar_erro

With txtDescricao_serv
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercialServ
    .Locked = False
    .TabStop = True
End With
With cmbFamilia_serv
    .Locked = False
    .TabStop = True
End With
With cmbUn_serv
    .Locked = False
    .TabStop = True
End With
'With Cmb_un_com_serv
'    .Locked = False
'    .TabStop = True
'End With
With txtValorUnit_serv
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosServ()
On Error GoTo tratar_erro

If txtCodigo <> "" Then
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select * from projproduto where desenho = '" & txtCodigo.Text & "' and Compras = 'True' and DtValidacao IS NOT NULL and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        txtcodproduto_serv = 0
        cmbUn_serv.ListIndex = -1
        Cmb_un_com_serv.ListIndex = -1
        cmbFamilia_serv.ListIndex = -1
        txtValorUnit_serv = "0,00000"
        txtDescricao_serv.Text = ""
        txtDescricao_comercialServ.Text = ""
        cmbreferencia_serv.Clear
        If TBCompras_Lista!Bloqueado = False Then
            txtcodproduto_serv = TBCompras_Lista!Codproduto
            txtCodigo = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
            txtDescricao_serv.Text = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
            txtDescricao_comercialServ.Text = IIf(IsNull(TBCompras_Lista!descricaotecnica), "", TBCompras_Lista!descricaotecnica)
            If IsNull(TBCompras_Lista!Unidade) = False And TBCompras_Lista!Unidade <> "" Then cmbUn_serv.Text = TBCompras_Lista!Unidade
            If IsNull(TBCompras_Lista!Unidade) = False And TBCompras_Lista!Unidade <> "" Then Cmb_un_com_serv.Text = TBCompras_Lista!Unidade
            'If IsNull(TBCompras_Lista!Unidade_com) = False And TBCompras_Lista!Unidade_com <> "" Then Cmb_un_com_serv.Text = TBCompras_Lista!Unidade_com
            If IsNull(TBCompras_Lista!Classe) = False And TBCompras_Lista!Classe <> "" Then cmbFamilia_serv.Text = TBCompras_Lista!Classe
2:
            valor = IIf(Txt_valor_moeda = "", 1, Txt_valor_moeda)
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Projproduto_fornecedor where Codproduto = " & TBCompras_Lista!Codproduto & " and idfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                txtValorUnit_serv = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto / valor, "###,##0.0000000000"))
            Else
                txtValorUnit_serv = IIf(IsNull(TBCompras_Lista!PCusto), "", (Format(TBCompras_Lista!PCusto / valor, "###,##0.0000000000")))
            End If
            TBFI.Close
            
            If TBCompras_Lista!Estoque = True Then
                With cmbUn_serv
                    .Locked = True
                    .TabStop = False
                End With
    '            With Cmb_un_com_serv
    '                .Locked = True
    '                .TabStop = False
    '            End With
            Else
                With cmbUn_serv
                    .Locked = False
                    .TabStop = True
                End With
    '            With Cmb_un_com_serv
    '                .Locked = False
    '                .TabStop = True
    '            End With
            End If
            
            ProcCarregaDadosCFOPProdServ IIf(IsNull(TBCompras_Lista!ID_CFOP), 0, TBCompras_Lista!ID_CFOP), False
            ProcCarregaComboCodRef cmbreferencia_serv, "P.desenho = '" & txtCodigo & "'", txtIDfornecedor, "F", True, True
            
            Txt_vlr_unit_ultima_compra_serv = FunVerifVlrUnitUltCompra(txtCodigo, 0)
        Else
            USMsgBox ("Não é permitido utilizar este serviço, pois o mesmo está bloqueado."), vbExclamation, "CAPRIND v5.0"
        End If
        ProcBloqueiaCamposServComCadastrado
    Else
        txtIDLista_serv = 0
        ProcLimpaCamposServ False
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposServSemCadastrado
    End If
    TBCompras_Lista.Close
Else
    txtIDLista_serv = 0
    ProcLimpaCamposServ True
    If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposServSemCadastrado
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a unidade ou familia desse registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaTabsProd()
On Error GoTo tratar_erro

txtNomenclatura.TabStop = False
cmbReferencia.TabStop = False
txtEspecificacoes.TabStop = False
txtDescricao_comercial.TabStop = False
txtObs.TabStop = False
cmbfamilia.TabStop = False
cmbun.TabStop = False
Cmb_un_com.TabStop = False
Cmb_CST_ICMS.TabStop = False
txtdetalheitem.TabStop = False
txtvalorunitario.TabStop = False
txtFrete.TabStop = False
txtSeguro.TabStop = False
txtAcessorias.TabStop = False
txtQuantidade.TabStop = False
txtIPI.TabStop = False
txtICMS.TabStop = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaTabsServ()
On Error GoTo tratar_erro

txtCodigo.TabStop = False
cmbreferencia_serv.TabStop = False
txtDescricao_serv.TabStop = False
txtDescricao_comercialServ.TabStop = False
txtObs_serv.TabStop = False
cmbFamilia_serv.TabStop = False
txtDetalhe_serv.TabStop = False
cmbUn_serv.TabStop = False
'Cmb_un_com_serv.TabStop = False
txtValorUnit_serv.TabStop = False
txtQtde_serv.TabStop = False
txtISSQN.TabStop = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaTabsProd()
On Error GoTo tratar_erro

txtNomenclatura.TabStop = True
cmbReferencia.TabStop = True
txtEspecificacoes.TabStop = True
txtDescricao_comercial.TabStop = True
txtObs.TabStop = True
cmbfamilia.TabStop = True
cmbun.TabStop = True
Cmb_un_com.TabStop = True
Cmb_CST_ICMS.TabStop = True
txtdetalheitem.TabStop = True
txtvalorunitario.TabStop = True
txtFrete.TabStop = True
txtSeguro.TabStop = True
txtAcessorias.TabStop = True
txtQuantidade.TabStop = True
txtIPI.TabStop = True
txtICMS.TabStop = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaTabsServ()
On Error GoTo tratar_erro

txtCodigo.TabStop = True
cmbreferencia_serv.TabStop = True
txtDescricao_serv.TabStop = True
txtDescricao_comercialServ.TabStop = True
txtObs_serv.TabStop = True
cmbFamilia_serv.TabStop = True
txtDetalhe_serv.TabStop = True
cmbUn_serv.TabStop = True
'Cmb_un_com_serv.TabStop = True
txtValorUnit_serv.TabStop = True
txtQtde_serv.TabStop = True
txtISSQN.TabStop = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_codigo_Click()
On Error GoTo tratar_erro

ProcLimpaCamposServ False
ProcCarregaDadosServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLimpar_CFOP_serv_Click()
On Error GoTo tratar_erro

Txt_ID_CFOP_serv = ""
txtCFOP_serv = ""
txtNatureza_operacao_serv = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLimpar_NCM_Click()
On Error GoTo tratar_erro

Txt_ID_CF = ""
Txt_CF = ""
ProcCalculaValor True
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalentrega_Click()
On Error GoTo tratar_erro

Txt_ID_entrega = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!endereco_entrega) = False And TBAbrir!endereco_entrega <> "" Then
        txtlocal.Text = TBAbrir!endereco_entrega
    Else
        If IsNull(TBAbrir!Tipo_endereco) = False And TBAbrir!Tipo_endereco <> "" Then
            Endereco = TBAbrir!Tipo_endereco & ": " & IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
        Else
            Endereco = IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
        End If
        If IsNull(TBAbrir!Tipo_bairro) = False And TBAbrir!Tipo_bairro <> "" Then
            Bairro = TBAbrir!Tipo_bairro & ": " & IIf(IsNull(TBAbrir!Bairro), "", TBAbrir!Bairro)
        Else
            Bairro = IIf(IsNull(TBAbrir!Bairro), "", TBAbrir!Bairro)
        End If
        txtlocal.Text = Endereco & " - " & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade) & " - " & IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF) & " - " & IIf(IsNull(TBAbrir!CEP), "", TBAbrir!CEP)
        Txt_ID_entrega = TBAbrir!CODIGO
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarEscopo()
On Error GoTo tratar_erro

Novo_PC3 = False
Aplic = 4
Compras_Cotacao = False
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocTransp_Click()
On Error GoTo tratar_erro



With Cmb_tipo_transp
    Acao = "localizar a transportadora"
    If .Text = "" Then
        NomeCampo = "o tipo da transportadora"
        ProcVerificaAcao
        .SetFocus
        Exit Sub
    End If
    
    Sit_REG = 3
    If .Text = "Cliente" Then
        ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
        frmVendas_LocalizarCliente.Show 1
    ElseIf .Text = "Fornecedor" Then
            ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
            Compras_Pedido = True
            FrmCompras_localizafornecedor.Show 1
        Else
            frmFaturamento_Prod_Serv_Localizar_Empresa.Show 1
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoTab_Servico()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab3.Tab
    Case 0: Nome_anexo = "serviço"
    Case 1: Nome_anexo = "centro de custo"
    Case 2: Nome_anexo = "empenho"
End Select
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", Nome_anexo, "criar novo", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", Nome_anexo, "criar novo", True, True) = False Then Exit Sub
End If
If SSTab2.Tab <> 2 Then If FunVerifSatus("criar novo " & Nome_anexo, True) = False Then Exit Sub

Select Case SSTab3.Tab
    Case 0:
        txtIDLista_serv = 0
        Novo_PC2 = True
        ProcLimpaCamposServ True
        Frame1(7).Enabled = True
        ProcDesbloqueiaCamposServ
        txtCodigo.SetFocus
        ProcLiberaTabsServ
    Case 1:
        ProcLimpaCamposCustoServ
        Frame1(8).Enabled = True
        Novo_PC2_Custo = True
        Cmb_centro_servico.SetFocus
    Case 2:
        Sit_REG = 1
        Compras_Requisicao = False
        Compras_Cotacao = False
        Compras_Pedido = True
        frmProd_Lista_Produto.Show 1
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoEscopo()
On Error GoTo tratar_erro

If txtStatus = "CANCELADO" Then
    USMsgBox ("Não é permitido criar novo escopo de fornecimento para este pedido de compra, pois o mesmo está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "escopo de fornecimento", "criar novo", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "escopo de fornecimento", "criar novo", True, True) = False Then Exit Sub
End If
If FunVerifSatus("criar novo escopo de fornecimento", True) = False Then Exit Sub
txtEscopo = ""
Novo_PC3 = True
Aplic = 4
Compras_Cotacao = False
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDPedido = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from compras_pedido order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("idpedido = " & txtIDPedido)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtIDPedido.Text = TBLISTA!IDpedido
        Set TBCompras_Pedido = CreateObject("adodb.recordset")
        TBCompras_Pedido.Open "Select * from compras_pedido where idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        ProcLimparComercial
        TXTIDLista = 0
        ProcLimpaCamposItem True
        ProcLimpaCamposCusto
        SSTab2.Tab = 0
        txtIDLista_serv = 0
        ProcLimpaCamposServ True
        ProcLimpaCamposCustoServ
        SSTab3.Tab = 0
        ProcPuxaDados
        ProcAbreComercial
        ProcCarregaEscopoForn
        ProcAtualizalista
        ProcAtualizalistaServ
    Else
        USMsgBox ("Fim dos cadastros de pedido de compra."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_PC1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarTab_Servico()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab3.Tab
    Case 0: ProcGravarServ
    Case 1: ProcGravarServ_custo
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarEscopo()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM compras_comercial WHERE IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If Novo_PC3 = True Then
        Evento = "Novo escopo de fornecimento"
        USMsgBox ("Novo escopo de fornecimento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        If txtResponsavel_aprovacao <> "" Then
            If txtResponsavel_aprovacao <> pubUsuario Then
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "escopo de fornecimento", "alterar", True, True) = False Then Exit Sub
            End If
        Else
            If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "escopo de fornecimento", "alterar", True, True) = False Then Exit Sub
        End If
        If FunVerifSatus("alterar o escopo de fornecimento", True) = False Then Exit Sub
        
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar escopo de fornecimento"
    End If
Else
    TBProduto.AddNew
    USMsgBox ("Novo escopo de fornecimento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
End If
'==================================
Modulo = "Compras/Pedido"
ID_documento = txtIDPedido
Documento = "Nº pedido: " & txtPedido.Text
Documento1 = ""
ProcGravaEvento
'==================================
TBProduto!IDpedido = txtIDPedido.Text
TBProduto!Escopo = txtEscopo
TBProduto.Update
TBProduto.Close
Novo_PC3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab4.Tab = 0 Then
            If TBLISTA_Pedido_Necessidade.AbsolutePage <> 2 Then
                If TBLISTA_Pedido_Necessidade.AbsolutePage = -3 Then
                    ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.PageCount - 1)
                Else
                    TBLISTA_Pedido_Necessidade.AbsolutePage = TBLISTA_Pedido_Necessidade.AbsolutePage - 2
                    ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.AbsolutePage)
                End If
            Else
                ProcExibePagina_Necessidade (1)
            End If
        Else
            If TBLISTA_Pedido_Solicitacao.AbsolutePage <> 2 Then
                If TBLISTA_Pedido_Solicitacao.AbsolutePage = -3 Then
                    ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.PageCount - 1)
                Else
                    TBLISTA_Pedido_Solicitacao.AbsolutePage = TBLISTA_Pedido_Solicitacao.AbsolutePage - 2
                    ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.AbsolutePage)
                End If
            Else
                ProcExibePagina_Solicitacao (1)
            End If
        End If
    Case 1:
        If TBLISTA_Compras_Pedido.AbsolutePage <> 2 Then
            If TBLISTA_Compras_Pedido.AbsolutePage = -3 Then
                ProcExibePagina (TBLISTA_Compras_Pedido.PageCount - 1)
            Else
                TBLISTA_Compras_Pedido.AbsolutePage = TBLISTA_Compras_Pedido.AbsolutePage - 2
                ProcExibePagina (TBLISTA_Compras_Pedido.AbsolutePage)
            End If
        Else
            ProcExibePagina (1)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4))
If Quant <= 1 Or txtPagIr(index) > Quant Then Exit Sub
If txtPagIr(index).Text >= 1 And txtPagIr(index).Text <= Quant Then
    Select Case SSTab1.Tab
        Case 0:
            If SSTab4.Tab = 0 Then
                TBLISTA_Pedido_Necessidade.AbsolutePage = txtPagIr(index).Text
                ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.AbsolutePage)
            Else
                TBLISTA_Pedido_Solicitacao.AbsolutePage = txtPagIr(index).Text
                ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.AbsolutePage)
            End If
        Case 1:
            TBLISTA_Compras_Pedido.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina (TBLISTA_Compras_Pedido.AbsolutePage)
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab4.Tab = 0 Then
            TBLISTA_Pedido_Necessidade.AbsolutePage = 1
            ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.AbsolutePage)
        Else
            TBLISTA_Pedido_Solicitacao.AbsolutePage = 1
            ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.AbsolutePage)
        End If
    Case 1:
        TBLISTA_Compras_Pedido.AbsolutePage = 1
        ProcExibePagina (TBLISTA_Compras_Pedido.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab4.Tab = 0 Then
            If TBLISTA_Pedido_Necessidade.AbsolutePage <> -3 Then
                If TBLISTA_Pedido_Necessidade.AbsolutePage = 1 Then
                    ProcExibePagina_Necessidade (2)
                Else
                    ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.AbsolutePage)
                End If
            Else
                ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.PageCount)
            End If
        Else
            If TBLISTA_Pedido_Solicitacao.AbsolutePage <> -3 Then
                If TBLISTA_Pedido_Solicitacao.AbsolutePage = 1 Then
                    ProcExibePagina_Solicitacao (2)
                Else
                    ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.AbsolutePage)
                End If
            Else
                ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.PageCount)
            End If
        End If
    Case 1:
        If TBLISTA_Compras_Pedido.AbsolutePage <> -3 Then
            If TBLISTA_Compras_Pedido.AbsolutePage = 1 Then
                ProcExibePagina (2)
            Else
                ProcExibePagina (TBLISTA_Compras_Pedido.AbsolutePage)
            End If
        Else
            ProcExibePagina (TBLISTA_Compras_Pedido.PageCount)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab4.Tab = 0 Then
            TBLISTA_Pedido_Necessidade.AbsolutePage = TBLISTA_Pedido_Necessidade.PageCount
            ProcExibePagina_Necessidade (TBLISTA_Pedido_Necessidade.AbsolutePage)
        Else
            TBLISTA_Pedido_Solicitacao.AbsolutePage = TBLISTA_Pedido_Solicitacao.PageCount
            ProcExibePagina_Solicitacao (TBLISTA_Pedido_Solicitacao.AbsolutePage)
        End If
    Case 1:
        TBLISTA_Compras_Pedido.AbsolutePage = TBLISTA_Compras_Pedido.PageCount
        ProcExibePagina (TBLISTA_Compras_Pedido.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvar_desconto_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "o valor total do desconto", "salvar", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "o valor total do desconto", "salvar", True, True) = False Then Exit Sub
End If

If FunVerifSatus("salvar o valor total do desconto", True) = False Then Exit Sub
Desconto = IIf(txtTotaldesconto = "", 0, txtTotaldesconto) 'Verifica valor do desconto
Acao = "salvar o desconto"
If Desconto < 0 Then
    NomeCampo = "o valor total do desconto"
    ProcVerificaAcao
    txtTotaldesconto.SetFocus
    Exit Sub
End If

vlrTotalProd = txt_vlrtotalprod
VlrTotalServ = txttotalservicos
If Desconto > (vlrTotalProd + VlrTotalServ) Then
    USMsgBox ("O valor total do desconto não pode ser maior que o valor total dos produtos + o valor total dos serviços."), vbExclamation, "CAPRIND v5.0"
    txtTotaldesconto = ""
    txtTotaldesconto.SetFocus
    Exit Sub
End If

'Valor Total de produtos e serviços
VlttTotal = 0
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select SUM(preco_unitario * Quant_Comp) as VlttTotal from Compras_pedido_lista where IDPedido = " & txtIDPedido & " and Remessa = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    VlttTotal = IIf(IsNull(TBProduto!VlttTotal), 0, TBProduto!VlttTotal)
End If
    
Contador = 0
valor = 0

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Compras_pedido_lista where IDPedido = " & txtIDPedido & " and Remessa = 'False' order by IdLista", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Do While TBProduto.EOF = False
        QuantsolicitadoN2 = 0 'Desconto
    
        VltUnit = Format(TBProduto!preco_unitario * TBProduto!Quant_Comp, "###,##0.00")
        If VlttTotal <> 0 Then Qtd = (VltUnit * 100) / VlttTotal
        
        Contador = Contador + 1
        
        If Contador = TBProduto.RecordCount Then
            QuantsolicitadoN2 = Desconto - valor 'Desconto
        Else
            If Desconto <> 0 Then QuantsolicitadoN2 = (Desconto * Qtd) / 100 'Desconto
        End If
        TBProduto!ValorDesconto = Format(QuantsolicitadoN2 / TBProduto!Quant_Comp, "###,##0.0000000000")
        TBProduto!Desconto = Format((TBProduto!ValorDesconto / TBProduto!preco_unitario) * 100, "###,##0.0000000000")
        TBProduto!preco_unitario_desconto = Format(TBProduto!preco_unitario - Round(TBProduto!ValorDesconto, 2), "###,##0.0000000000")
        
        Valor1 = Format(TBProduto!preco_unitario_desconto * TBProduto!Quant_Comp, "###,##0.00") + IIf(IsNull(TBProduto!Frete), 0, TBProduto!Frete)
        TBProduto!VlrISSQN = Format((Valor1 * IIf(IsNull(TBProduto!ISSQN), 0, TBProduto!ISSQN)) / 100, "###,##0.00")
        If IsNull(TBProduto!ID_CF) = True Or TBProduto!ID_CF = "" Or TBProduto!ID_CF = "0" Then
            IntIPI = IIf(IsNull(TBProduto!IPI), 0, TBProduto!IPI)
            IntICMS = IIf(IsNull(TBProduto!ICMS), 0, TBProduto!ICMS)
            
            VlrIPI = Format((Valor1 * IntIPI) / 100, "###,##0.00") 'Calcula IPI
            TBProduto!BC_ICMS = Valor1
            VlrICMS_suframa = Format((Valor1 * IntICMS) / 100, "###,##0.00") 'Calcula ICMS
        Else
            ProcValorImposto txtPedido, TBProduto!ID_CF, IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), 0
            ProcControleImposto IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), 0
            If TemIPI = "NÃO" Then IntIPI = 0
            If TemICMS = "NÃO" Then IntICMS = 0
            
            VlrIPI = Format((Valor1 * IntIPI) / 100, "###,##0.00") 'Calcula IPI
            'Calclula ICMS
            ProcCalculaBC Cmb_empresa.ItemData(Cmb_empresa.ListIndex), IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), 0, Valor1, VlrIPI, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBProduto!CST), 0, TBProduto!CST), "T", txtIDfornecedor, txtFornecedor
            TBProduto!BC_ICMS = BC
            VlrICMS_suframa = Format((BC * IntICMS) / 100, "###,##0.00")
            
            TBProduto!Valor_ICMS_ST = 0
            TBProduto!BC_ICMS_ST = 0
            If IsNull(TBProduto!CST) = False And TBProduto!CST <> "" And IsNull(TBProduto!ID_CFOP) = False And TBProduto!ID_CFOP <> "" Then
                Set TBCFOP = CreateObject("adodb.recordset")
                TBCFOP.Open "Select id_CFOP from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    ProcSubstituicaoTributaria txtuf, TBProduto!CST, TBProduto!ID_CF, IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, TBProduto!preco_unitario_desconto, TBProduto!Quant_Comp, TBProduto!ValorDesconto * TBProduto!Quant_Comp, BC, 0, 0, 0, False, False, 0
                    TBProduto!Valor_ICMS_ST = ICMSCST
                    If ICMSCST <> 0 Then TBProduto!BC_ICMS_ST = BCICMSCST
                End If
                TBCFOP.Close
            End If
        End If
        TBProduto!VlrIPI = VlrIPI
        TBProduto!vlrICMS = VlrICMS_suframa
        TBProduto!preco_total = Format(Valor1, "###,##0.00")
        TBProduto.Update
        
        valor = valor + QuantsolicitadoN2
        TBProduto.MoveNext
    Loop
End If
TBProduto.Close
ProcAtualizalista
ProcAtualizalistaServ
ProcGravarTotaisPC txtIDPedido

USMsgBox ("Valor total do desconto salvo com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Compras/Pedido"
Evento = "Salvar valor total do desconto no pedido"
ID_documento = txtIDPedido
Documento = "Nº pedido: " & txtPedido.Text
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvar_Frete_Click()
On Error GoTo tratar_erro
Dim Acessorias As Double

            
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "o valor total do frete, seguro e despesas acessórias", "salvar", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "o valor total do frete, seguro e despesas acessórias", "salvar", True, True) = False Then Exit Sub
End If
If FunVerifSatus("salvar o valor total do frete, seguro e despesas acessórias", True) = False Then Exit Sub

Acao = "salvar o frete/seguro/despesas acessórias"
valor = IIf(TxtTotalFrete = "", 0, TxtTotalFrete) 'Verifica valor do frete
If valor < 0 Then
    NomeCampo = "o valor total do frete"
    ProcVerificaAcao
    TxtTotalFrete.SetFocus
    Exit Sub
End If
valor = IIf(txtTotalSeguro = "", 0, txtTotalSeguro) 'Verifica valor do seguro
If valor < 0 Then
    NomeCampo = "o valor total do seguro"
    ProcVerificaAcao
    txtTotalSeguro.SetFocus
    Exit Sub
End If
valor = IIf(TxtTotalacessorias = "", 0, TxtTotalacessorias) 'Verifica valor de despesas acessórias
If valor < 0 Then
    NomeCampo = "o valor total das despesas acessórias"
    ProcVerificaAcao
    TxtTotalacessorias.SetFocus
    Exit Sub
End If

If IIf(TxtTotalFrete = "", 0, TxtTotalFrete) > 0 Then
    Conexao.Execute "Update compras_pedido_lista Set VlrIPI = VlrIPI - ((Frete * IPI) / 100) where IDPedido = " & txtIDPedido & " and Frete_IPI = 'True'"
    
    If USMsgBox("O valor do frete tem IPI?", vbYesNo, "CAPRIND v5.0") = vbYes Then TextoFiltro = "Frete_IPI = 'True'" Else TextoFiltro = "Frete_IPI = 'False'"
    Conexao.Execute "UPDATE compras_pedido_lista Set " & TextoFiltro & " where IDPedido = " & txtIDPedido
End If

ValorPago = IIf(TxtTotalFrete = "", 0, TxtTotalFrete) 'Verifica valor do frete
Seguro1 = IIf(txtTotalSeguro = "", 0, txtTotalSeguro) 'Verifica valor de seguro
Acessorias1 = IIf(TxtTotalacessorias = "", 0, TxtTotalacessorias) 'Verifica valor de acessorias

'Valor Total de produtos
VlttTotal = 0
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select SUM(preco_total) as ValorProdutos from compras_pedido_lista where IDPedido = " & txtIDPedido & " and Remessa = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    VlttTotal = IIf(IsNull(TBProduto!ValorProdutos), 0, TBProduto!ValorProdutos)
End If

Contador = 0
valor = 0
Valor1 = 0
Valor2 = 0
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from compras_pedido_lista where IDPedido = " & txtIDPedido & " and Remessa = 'False' order by IDLista", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Do While TBProduto.EOF = False
        Frete = 0
        Seguro = 0
        Acessorias = 0
    
        'Verifica valores para somar na base de calculo
        VltUnit = IIf(IsNull(TBProduto!preco_total), 0, TBProduto!preco_total)
        If VlttTotal <> 0 Then Qtd = (VltUnit * 100) / VlttTotal
        
        Contador = Contador + 1
        If Contador = TBProduto.RecordCount Then
            Frete = ValorPago - valor 'Frete
            Seguro = Seguro1 - Valor1 'Seguro
            Acessorias = Acessorias1 - Valor2 'Acessorias
        Else
            If ValorPago <> 0 Then Frete = Format((ValorPago * Qtd) / 100, "###,##0.00") 'Frete
            If Seguro1 <> 0 Then Seguro = Format((Seguro1 * Qtd) / 100, "###,##0.00") 'Seguro
            If Acessorias1 <> 0 Then Acessorias = (Acessorias1 * Qtd) / 100 'Acessorias
        End If
        TBProduto!Frete = Format(Frete, "###,##0.00")
        TBProduto!Seguro = Format(Seguro, "###,##0.00")
        TBProduto!Acessorias = Acessorias
        
        If TBProduto!Frete_IPI = True And TBProduto!Frete_IPI <> 0 Then
            TBProduto!VlrIPI = TBProduto!VlrIPI + ((Frete * IIf(IsNull(TBProduto!IPI), 0, TBProduto!IPI)) / 100)
        End If
        
        TBProduto.Update
        
        valor = valor + Frete
        Valor1 = Valor1 + Seguro
        Valor2 = Valor2 + Acessorias
        TBProduto.MoveNext
    Loop
End If
TBProduto.Close

ProcAtualizalista
ProcAtualizalistaServ
ProcGravarTotaisPC txtIDPedido

USMsgBox ("Valor total do frete, seguro e despesas acessórias salvos com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Compras/Pedido"
Evento = "Salvar valor total do frete, seguro e despesas acessórias"
ID_documento = txtIDPedido
Documento = "Nº pedido: " & txtPedido.Text
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro

'If txtPedido = "" Or (txtStatus <> "COMPRADO" And txtStatus <> "APROVADO") Then Exit Sub
'frmCompras_Pedido_Status.Show 1

If txtDtValidacao.Text = "" Then
  USMsgBox "Para sincronizar status do pedido de compras, é necessário validar primeiro!", vbInformation, "CAPRIND v5.0"
  Exit Sub
End If

If USMsgBox("Deseja realmente sincronizar o status do pedido de compras?", vbYesNo, "CAPRIND  v5.0") = vbYes Then
  If txtStatus.Text = "AGUARDANDO APROVAÇÃO" And txtDtValidacao.Text <> "" And txtIDPedido <> 0 Then
    ProcBuscarPedidoWEB (txtIDPedido)
    'ProcSalvarPedidoWEB (Int(txtIDPedido))
  End If
End If

  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdTransporte_padrao_Click()
On Error GoTo tratar_erro

Aplic = 5
Compras_Cotacao = False
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLimpar_CFOP_prod_Click()
On Error GoTo tratar_erro

Txt_ID_CFOP_prod = ""
txtCFOP_prod = ""
Txt_natureza_operacao_prod = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyF2: If SSTab4.Tab = 0 Then ProcFiltrar_Necessidade Else ProcFiltrar_Solicitacao
            Case vbKeyF3: ProcGerarPed
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiar
            Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros listapedido, "Compras/Pedido"
            Case vbKeyF10: If Cmb_opcao_lista = "Aprovação" Then ProcValidarRegistros listapedido, "Compras/Pedido/Aprovar"
            Case vbKeyF11: ProcEnviarEmail
            Case vbKeyF12: ProcExportarExcel
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF3: ProcSalvarComercial
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcFinanceiro
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoTab_Produto
            Case vbKeyF3: ProcSalvarTab_Produto
            Case vbKeyF4: ProcExcluirTab_Produto
            Case vbKeyF5: ProcImprimir
            Case vbKeyF8: If USToolBar4.ButtonState(7) = 0 Then procCalculadora
            Case vbKeyF9: If USToolBar4.ButtonState(8) = 0 Then ProcAlterarStatusItem
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoTab_Servico
            Case vbKeyF3: ProcSalvarTab_Servico
            Case vbKeyF4: ProcExcluirTab_Servico
            Case vbKeyF5: ProcImprimir
            Case vbKeyF9: If USToolBar5.ButtonState(7) = 0 Then ProcAlterarStatusServico
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 5:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoEscopo
            Case vbKeyF2: ProcLocalizarEscopo
            Case vbKeyF3: ProcSalvarEscopo
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Private Sub cmdAdicionarfornecedor_Click()
On Error GoTo tratar_erro

Sit_REG = 2
ProcConfVariaveisLocForn False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarTab_Produto()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0: ProcgravarItem
    Case 1: ProcGravarItem_custo
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcgravarItem()
On Error GoTo tratar_erro

If Frame1(12).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If

If LiberarData = True Then
    If USMsgBox("Deseja realmente alterar o prazo de entrega desse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        StrSql = "update compras_pedido_Lista set prazo = '" & txtprazo_item.Text & "' where idLista = '" & TXTIDLista.Text & "'"
        'Debug.print StrSql
        Conexao.Execute StrSql
        USMsgBox "Prazo de entrega alterado com sucesso!", vbInformation, "CAPRIND v5.0"
        If USMsgBox("Deseja colocar esse prazo em todos os itens do pedido?", vbYesNo, "CAPRIND v5.0") = vbYes Then
         StrSql = "update compras_pedido_Lista set prazo = '" & txtprazo_item.Text & "' where idPedido = '" & txtIDPedido & "' and status_item <> 'RECEBIDO' and status_item <> 'PARCIAL'"
         Debug.Print StrSql
         Conexao.Execute StrSql
         USMsgBox "Prazo de entrega alterado em todos os itens do pedido com sucesso!", vbInformation, "CAPRIND v5.0"
        End If
        LiberarData = False
        Exit Sub
    End If

End If


If Liberado = False Then
If txtstatus_item = "RECEBIDO" Then
USMsgBox "Não é permitido alterar cadastro de produto com status RECEBIDO!", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

If txtstatus_item = "RECEBIDO PARCIAL" Then
USMsgBox "Não é permitido alterar cadastro de produto com status RECEBIDO PARCIAL!", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

If txtstatus_item = "COMPRADO" Then
USMsgBox "Não é permitido alterar cadastro de produto com status COMPRADO!", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

If txtRespValidacao <> "" Then
USMsgBox "Não é permitido alterar cadastro de produto com pedido de compra validado!", vbCritical, "CAPRIND v5.0"
Exit Sub
End If
End If

Acao = "salvar"
If chkAuto.Value = 0 And txtNomenclatura = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtNomenclatura.SetFocus
    Exit Sub
End If
If txtEspecificacoes = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtEspecificacoes.SetFocus
    Exit Sub
End If
If txtDescricao_comercial = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtDescricao_comercial.SetFocus
    Exit Sub
End If
If IsDate(txtprazo_item) = False Then
    NomeCampo = "o prazo"
    ProcVerificaAcao
    txtprazo_item.SetFocus
    Exit Sub
End If
If cmbfamilia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If cmbun.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
If txtOrdem <> "" And txtOrdem <> "0" Then
    If FunVerifOPCarregaOS(Cmb_OS, txtOrdem, True, False) = False Then
        txtOrdem.SetFocus
        Exit Sub
    End If
End If
valor = IIf(txtvalorunitario = "", 0, txtvalorunitario)
If valor < 0 Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    txtvalorunitario.SetFocus
    Exit Sub
End If
If Chk_desc.Value = 1 Then
    valor = IIf(txtDesconto = "", 0, txtDesconto)
    If valor < 0 Or valor > 100 Then
        NomeCampo = "a porcentagem do desconto"
        ProcVerificaAcao
        txtDesconto.SetFocus
        Exit Sub
    End If
End If
If Chk_valor_desc.Value = 1 Then
    valor = IIf(txtvalordesconto = "", 0, txtvalordesconto)
    If valor < 0 Then
        NomeCampo = "o valor do desconto"
        ProcVerificaAcao
        txtvalordesconto.SetFocus
        Exit Sub
    End If
End If
valor = IIf(txtQuantidade = "", 0, txtQuantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQuantidade.SetFocus
    Exit Sub
End If
If txtIPI.Text = "" Then
    NomeCampo = "a porcentagem do IPI"
    ProcVerificaAcao
    txtIPI.SetFocus
    Exit Sub
End If
If txtICMS.Text = "" Then
    NomeCampo = "a porcentagem do ICMS"
    ProcVerificaAcao
    txtICMS.SetFocus
    Exit Sub
End If

If chkRemessa.Value = 0 Then
    If txtOrdem <> "" And txtOrdem <> "0" Then TextoFiltro = "and CPL.ordem = " & txtOrdem Else TextoFiltro = ""
   
    valor = txtvalorunitario
    NovoValor = Replace(valor, ",", ".")
'    Set TBCompras_Lista = CreateObject("adodb.recordset")
'    TBCompras_Lista.Open "Select CPL.IDlista from compras_pedido_lista CPL INNER JOIN projproduto P on CPL.Codproduto = P.codproduto where CPL.IDPedido = " & Txtidpedido & " and CPL.IDLista <> " & TXTIDLista & " and CPL.Desenho = '" & txtNomenclatura & "' and CPL.detalheitem = '" & txtdetalheitem & "' and CPL.Prazo = '" & Format(txtprazo_item, "Short Date") & "' " & TextoFiltro & " and CPL.Remessa = 'False' and CPL.preco_unitario = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
'    If TBCompras_Lista.EOF = False Then
'        USMsgBox "Não é permitido salvar este produto, pois o mesmo já foi cadastrado com esse valor unitário, prazo de entrega e detalhe.", vbExclamation, "CAPRIND v5.0"
'        If txtprazo_item.Enabled = True Then txtprazo_item.SetFocus
'        TBCompras_Lista.Close
'        Exit Sub
'    End If
'    TBCompras_Lista.Close
    
    'Verifica se o produto adicionado está requisitado na ordem ou é similar ao produto requisitado na ordem
    If txtOrdem <> "" And txtOrdem <> "0" And Cmb_OS = "" Then
        If FunVerifProdSimiliar(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Idmateriaprima from Producaomaterial where Ordem = " & txtOrdem, Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select Idmateriaprima from Producaomaterial where Ordem = " & txtOrdem & " and Codigo = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = True Then
                    
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select P.Codproduto from (Producaomaterial PM INNER JOIN projproduto P ON PM.Codigo = P.Desenho) INNER JOIN Projproduto P1 ON P1.ID_similar = P.ID_similar where PM.Ordem = " & txtOrdem & " and P1.Desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = True Then
                        USMsgBox "Não é permitido salvar este produto, pois o mesmo não é similar ao produto requisitado para esta ordem.", vbExclamation, "CAPRIND v5.0"
                        TBFIltro.Close
                        Exit Sub
                    End If
                End If
            End If
            TBFIltro.Close
        End If
    End If
    
    'Verifica se é obrigatório ter cotação valida
    If FunVerifCotacaoValida(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), txtNomenclatura, True, True, "salvar", txtIDfornecedor) = False Then Exit Sub
Else
    'Verifica se existe empenho para este produto e não deixa salvar como remessa
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ID from Compras_pedido_lista_empenhos where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        USMsgBox ("Não é permitido salvar este produto como remessa, pois o mesmo está empenhado."), vbExclamation, "CAPRIND v5.0"
        TBFIltro.Close
        Exit Sub
    End If
    TBFIltro.Close
End If

If chkAuto.Value = 1 Then
    ProcNovoProdutoAuto
    If txtreferencia <> "" Then
        cmbReferencia.AddItem txtreferencia
        cmbReferencia = txtreferencia
    End If
    chkAuto.Value = 0
End If
If chkManual.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtNomenclatura.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtNomenclatura.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManual
    If txtreferencia <> "" Then
        cmbReferencia.AddItem txtreferencia
        cmbReferencia = txtreferencia
    End If
    chkManual.Value = 0
End If

'Verifica se o produto está cadastrado
If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtNomenclatura.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then
        USMsgBox ("Não é permitido salvar este produto, pois o mesmo não está cadastrado."), vbExclamation, "CAPRIND v5.0"
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If

'Se a unidade for diferente verifica se esta cadastrado o peso bruto e UN/KG
If FunBloqueiaUnConversao(txtNomenclatura, cmbun, Cmb_un_com, True) = True Then Exit Sub

'=============================================================================
'Se a data de entrega for menor que a data atual pergunta se deseja gravar
'=============================================================================
If IsDate(txtprazo_item) Then
PrazoEntrega = txtprazo_item
    If PrazoEntrega < Date Then
        If USMsgBox("O prazo de entrega menor que a data atual, deseja continuar?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            Exit Sub
        End If
    End If
End If
'=============================================================================

Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from compras_pedido_lista where idlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = True Then
    TBCompras_Lista.AddNew
    TBCompras_Lista!Status_Item = IIf(txtstatus_item = "COMPRADO", "N_RECEBIDO", txtstatus_item)
Else
If Liberado = False Then
    If txtResponsavel_aprovacao <> "" Then
        If txtResponsavel_aprovacao <> pubUsuario Then
            If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este produto", "alterar", True, True) = False Then Exit Sub
        End If
    Else
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este produto", "alterar", True, True) = False Then Exit Sub
    End If
'    If FunVerifSatus("alterar este produto", True) = False Then Exit Sub
'    If FunVerifSatusProdServ(txtstatus_item, "alterar este produto", True, True) = False Then Exit Sub
End If
End If
TBCompras_Lista!IDpedido = txtIDPedido
If TBCompras_Lista!Remessa = False And (chkRemessa.Value = 0 Or chkRemessa.Value = 1) Then Permitido = True Else Permitido = False
If TBCompras_Lista!Status_Item <> "RECEBIDO" Then
    TBCompras_Lista!Codproduto = txtcodproduto
    TBCompras_Lista!Desenho = txtNomenclatura.Text
    TBCompras_Lista!N_referencia = cmbReferencia
    TBCompras_Lista!Descricao = txtEspecificacoes.Text
    TBCompras_Lista!Descricao_comercial = txtDescricao_comercial.Text
    TBCompras_Lista!detalheitem = txtdetalheitem.Text
    TBCompras_Lista!Quant_Comp = txtQuantidade.Text
    TBCompras_Lista!Quant_Comp_PC = IIf(txtQuantidade_PC = "", Null, txtQuantidade_PC)
    TBCompras_Lista!Desconto = IIf(txtDesconto.Text = "", 0, txtDesconto.Text)
    TBCompras_Lista!ValorDesconto = IIf(txtvalordesconto.Text = "", 0, txtvalordesconto.Text)
    TBCompras_Lista!preco_unitario_desconto = txtvalorunitariodesc
    TBCompras_Lista!preco_unitario = txtvalorunitario
    TBCompras_Lista!preco_total = txtvlrTotal.Text
    TBCompras_Lista!Familia = cmbfamilia.Text
    TBCompras_Lista!Un = cmbun.Text
    TBCompras_Lista!Unidade_com = Cmb_un_com.Text
    TBCompras_Lista!IPI = IIf(txtIPI = "", Null, txtIPI)
    TBCompras_Lista!VlrIPI = TxtvlrIpi
    TBCompras_Lista!ICMS = IIf(txtICMS = "", Null, txtICMS)
    TBCompras_Lista!vlrICMS = txtvlrICMS
    
    'Cadastrar automaticamente a alteração da data
    If Novo_PC1 = False And IsNull(TBCompras_Lista!Prazo) = False Then
        If txtprazo_item <> TBCompras_Lista!Prazo Then ProcINSERTINTO "vendas_carteira_alteracoes", "ID_carteira, Data, Responsavel, Data_alteracao, Responsavel_alteracao, Obs, Alteracao_prazo, Padrao, Tipo", "" & TXTIDLista & ",'" & Date & "','" & pubUsuario & "','" & Date & "','" & pubUsuario & "', NULL,'ALTERADO O PRAZO DE ENTREGA DE " & Format(TBCompras_Lista!Prazo, "dd/mm/yy") & " PARA " & Format(txtprazo_item, "dd/mm/yy") & "', 'True', 'CPE'"
    End If
    TBCompras_Lista!Prazo = txtprazo_item
    
    TBCompras_Lista!Obs_pedido = IIf(txtObs = "", Null, txtObs)
    TBCompras_Lista!Tipo = "P"
    TBCompras_Lista!ID_CFOP = IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod)
    TBCompras_Lista!ID_CF = IIf(Txt_ID_CF = "", 0, Txt_ID_CF)
    TBCompras_Lista!CST = Cmb_CST_ICMS
    TBCompras_Lista!Frete = IIf(txtFrete = "", 0, txtFrete)
    TBCompras_Lista!Seguro = IIf(txtSeguro = "", 0, txtSeguro)
    TBCompras_Lista!Acessorias = IIf(txtAcessorias = "", 0, txtAcessorias)
    If ChkFrete_IPI.Value = 1 Then TBCompras_Lista!Frete_IPI = True Else TBCompras_Lista!Frete_IPI = False
'=============================================================================================
'Calcula base de cálculo do icms do item - PLE 28-10-2019
'=============================================================================================
    If Cmb_CST_ICMS <> "" And chkRemessa = 0 Then
     ProcCalculaBC Cmb_empresa.ItemData(Cmb_empresa.ListIndex), IIf(Txt_CFOP_prod = "", "0.000", Txt_CFOP_prod), 0, (TBCompras_Lista!preco_unitario_desconto * TBCompras_Lista!Quant_Comp) + TBCompras_Lista!Frete, TxtvlrIpi, SomarIPI, SomarIPIST, TemReducaoBC, False, Cmb_CST_ICMS, "T", txtIDfornecedor, txtFornecedor
    Else
     BC = 0
    End If

    TBCompras_Lista!BC_ICMS = BC
'=============================================================================================
' Calcula base de calculo do icms St do Item - PLE 28-10-2019
'=============================================================================================
    If Txt_ID_CF <> "" Then
        If Cmb_CST_ICMS <> "" And chkRemessa = 0 Then
            ProcSubstituicaoTributaria txtuf, Cmb_CST_ICMS, Txt_ID_CF, IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtvalorunitariodesc, txtQuantidade, BC, BCST, 0, 0, 0, False, False, 0
            TBCompras_Lista!Valor_ICMS_ST = ICMSCST
            If ICMSCST <> 0 Then TBCompras_Lista!BC_ICMS_ST = BCICMSCST Else TBCompras_Lista!BC_ICMS_ST = 0
        Else
            TBCompras_Lista!Valor_ICMS_ST = 0
            TBCompras_Lista!BC_ICMS_ST = 0
        End If
    End If
'=============================================================================================
    If chkRemessa.Value = 1 Then TBCompras_Lista!Remessa = True Else TBCompras_Lista!Remessa = False
    
    'Calcula quantidade se a unidade for diferente
    If cmbun <> Cmb_un_com Then
        If FunVerifUNConversao(cmbun, Cmb_un_com) = True Then
            TBCompras_Lista!Qtde_estoque = FunConverteUN(cmbun, Cmb_un_com, txtQuantidade, txtNomenclatura)
        Else
            TBCompras_Lista!Qtde_estoque = txtQuantidade / FunVerificaTabelaConversaoUnidade(cmbun, Cmb_un_com)
        End If
    Else
        TBCompras_Lista!Qtde_estoque = Null
    End If
End If
TBCompras_Lista.Update
TXTIDLista = TBCompras_Lista!IDlista

OF = IIf(txtOrdem = "", 0, txtOrdem)
If Cmb_OS = "" Then TextoFiltro = "OS = Null" Else TextoFiltro = "OS = " & Cmb_OS
Conexao.Execute "Update NFP Set NFP.Ordem = " & OF & " from tbl_Detalhes_Nota NFP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFPP.ID_carteira = " & TBCompras_Lista!IDlista & " and NFPP.Codinterno = '" & txtNomenclatura & "'"
Conexao.Execute "Update compras_pedido_lista Set Ordem = " & OF & ", " & TextoFiltro & " where idlista = " & TXTIDLista

ProcAgregarProdutoForn txtcodproduto, txtIDfornecedor, txtvalorunitario

ProcAtualizalista
If Novo_PC1 = True Then
    USMsgBox ("Novo produto cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto"
    If CodigoLista1 <> 0 And Listprod.ListItems.Count <> 0 Then
        Listprod.SelectedItem = Listprod.ListItems(CodigoLista1)
        Listprod.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Pedido"
ID_documento = TXTIDLista
Documento = "Nº pedido: " & txtPedido
Documento1 = "Cód. interno: " & txtNomenclatura
ProcGravaEvento
'==================================
Novo_PC1 = False

'Recalcula centro de custo
valor = txtvlrTotal
NovoValor = Replace(valor, ",", ".")
Conexao.Execute "Update compras_pedido_lista_custo Set valor = (ISNULL(Percentual, 0) * " & NovoValor & ") / 100 where IDlista = " & TXTIDLista
Liberado = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarServ()
On Error GoTo tratar_erro

If LiberarData = True Then
    If USMsgBox("Deseja realmente alterar o prazo de entrega desse serviço?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        StrSql = "update compras_pedido_Lista set prazo = '" & txtPrazo_serv & "' where idLista = '" & IDlista & "'"
        'Debug.print StrSql
        Conexao.Execute StrSql
        USMsgBox "Prazo de entrega alterado com sucesso!", vbInformation, "CAPRIND v5.0"
        If USMsgBox("Deseja colocar esse prazo em todos os serviços do pedido?", vbYesNo, "CAPRIND v5.0") = vbYes Then
         StrSql = "update compras_pedido_Lista set prazo = '" & txtPrazo_serv & "' where idPedido = '" & txtIDPedido & "'"
         Conexao.Execute StrSql
         USMsgBox "Prazo de entrega alterado em todos os serviços do pedido com sucesso!", vbInformation, "CAPRIND v5.0"
        End If
        LiberarData = False
        Exit Sub
    End If
End If


If txtStatus_serv.Text = "APROVADO" Then
    USMsgBox "Não é permitido alterar dados do serviço com status como APROVADO!", vbCritical, "CAPRIND v5.0"
    Exit Sub
End If

If Frame1(7).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If


Acao = "salvar"
If chkAuto_serv.Value = 0 And txtCodigo = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodigo.SetFocus
    Exit Sub
End If
If txtDescricao_serv = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricao_serv.SetFocus
    Exit Sub
End If
If txtDescricao_comercialServ = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtDescricao_comercialServ.SetFocus
    Exit Sub
End If
If IsDate(txtPrazo_serv) = False Then
    NomeCampo = "o prazo"
    ProcVerificaAcao
    txtPrazo_serv.SetFocus
    Exit Sub
End If
If cmbFamilia_serv.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbFamilia_serv.SetFocus
    Exit Sub
End If
If txtOrdem_serv <> "" And txtOrdem_serv <> "0" Then
    If FunVerifOPCarregaOS(Cmb_OS_serv, txtOrdem_serv, True, False) = False Then
        txtOrdem_serv.SetFocus
        Exit Sub
    End If
    If Cmb_OS_serv = "" Then
        NomeCampo = "o número da OS"
        ProcVerificaAcao
        Cmb_OS_serv.SetFocus
        Exit Sub
    End If
End If
If cmbUn_serv.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbUn_serv.SetFocus
    Exit Sub
End If
If Cmb_un_com_serv.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com_serv.SetFocus
    Exit Sub
End If
valor = IIf(txtValorUnit_serv = "", 0, txtValorUnit_serv)
If valor < 0 Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    txtValorUnit_serv.SetFocus
    Exit Sub
End If
If Chk_desc2.Value = 1 Then
    valor = IIf(txtDesconto_serv = "", 0, txtDesconto_serv)
    If valor < 0 Or valor > 100 Then
        NomeCampo = "a porcentagem do desconto"
        ProcVerificaAcao
        txtDesconto_serv.SetFocus
        Exit Sub
    End If
End If
If Chk_valor_desc2.Value = 1 Then
    valor = IIf(txtVlrDesconto_serv = "", 0, txtVlrDesconto_serv)
    If valor < 0 Then
        NomeCampo = "o valor do desconto"
        ProcVerificaAcao
        txtVlrDesconto_serv.SetFocus
        Exit Sub
    End If
End If
valor = IIf(txtQtde_serv = "", 0, txtQtde_serv)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQtde_serv.SetFocus
    Exit Sub
End If
If txtISSQN.Text = "" Then txtISSQN = 0

If chkAuto_serv.Value = 1 Then
    ProcnovoServicoAuto
    If txtReferencia_serv <> "" Then
        cmbreferencia_serv.AddItem txtReferencia_serv
        cmbreferencia_serv = txtReferencia_serv
    End If
    chkAuto_serv.Value = 0
End If
If chkManual_serv.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um serviço cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtCodigo.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoServicoManual
    If txtReferencia_serv <> "" Then
        cmbreferencia_serv.AddItem txtReferencia_serv
        cmbreferencia_serv = txtReferencia_serv
    End If
    chkManual_serv.Value = 0
End If

'Verifica se o serviço está cadastrado
If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtCodigo.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then
        USMsgBox ("Não é permitido salvar este serviço, pois o mesmo não está cadastrado."), vbExclamation, "CAPRIND v5.0"
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If
'Verifica se é obrigatório ter cotação valida
If FunVerifCotacaoValida(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), txtCodigo, False, True, "salvar", txtIDfornecedor) = False Then Exit Sub

'Se a unidade for diferente verifica se esta cadastrado o peso bruto e UN/KG
If FunBloqueiaUnConversao(txtCodigo, cmbUn_serv, Cmb_un_com_serv, False) = True Then Exit Sub

Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from compras_pedido_lista where idlista = " & txtIDLista_serv, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = True Then
    TBCompras_Lista.AddNew
    TBCompras_Lista!Status_Item = IIf(txtStatus_serv = "COMPRADO", "N_RECEBIDO", txtStatus_serv)
Else
If Liberado = False Then
    If txtResponsavel_aprovacao <> "" Then
        If txtResponsavel_aprovacao <> pubUsuario Then
            If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este serviço", "alterar", True, True) = False Then Exit Sub
        End If
    Else
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este serviço", "alterar", True, True) = False Then Exit Sub
    End If
    If FunVerifSatus("alterar este serviço", True) = False Then Exit Sub
    
    If FunVerifSatusProdServ(txtStatus_serv, "alterar este serviço", True, False) = False Then Exit Sub
End If
End If
TBCompras_Lista!IDpedido = txtIDPedido
If TBCompras_Lista!Status_Item <> "RECEBIDO" Or Liberado = True Then
    TBCompras_Lista!Codproduto = txtcodproduto_serv
    TBCompras_Lista!Desenho = txtCodigo
    TBCompras_Lista!N_referencia = cmbreferencia_serv
    TBCompras_Lista!Descricao = txtDescricao_serv
    TBCompras_Lista!Descricao_comercial = txtDescricao_comercialServ
    TBCompras_Lista!detalheitem = txtDetalhe_serv.Text
    TBCompras_Lista!Quant_Comp = txtQtde_serv.Text
    TBCompras_Lista!Desconto = IIf(txtDesconto_serv.Text = "", 0, txtDesconto_serv.Text)
    TBCompras_Lista!ValorDesconto = IIf(txtVlrDesconto_serv.Text = "", 0, txtVlrDesconto_serv.Text)
    TBCompras_Lista!preco_unitario_desconto = txtVlrUnitDesc_serv
    TBCompras_Lista!preco_unitario = txtValorUnit_serv
    TBCompras_Lista!preco_total = txtValorTotal_serv.Text
    TBCompras_Lista!ISSQN = txtISSQN.Text
    TBCompras_Lista!VlrISSQN = txtValor_ISSQN
    TBCompras_Lista!Familia = cmbFamilia_serv.Text
    TBCompras_Lista!Un = cmbUn_serv.Text
    TBCompras_Lista!Unidade_com = Cmb_un_com_serv.Text
    
    'Cadastrar automaticamente a alteração da data
    If Novo_PC2 = False And IsNull(TBCompras_Lista!Prazo) = False Then
        If txtPrazo_serv <> TBCompras_Lista!Prazo Then ProcINSERTINTO "vendas_carteira_alteracoes", "ID_carteira, Data, Responsavel, Data_alteracao, Responsavel_alteracao, Obs, Alteracao_prazo, Padrao, Tipo", "" & txtIDLista_serv & ",'" & Date & "','" & pubUsuario & "','" & Date & "','" & pubUsuario & "', NULL,'ALTERADO O PRAZO DE ENTREGA DE " & Format(TBCompras_Lista!Prazo, "dd/mm/yy") & " PARA " & Format(txtPrazo_serv, "dd/mm/yy") & "', 'True', 'CPE'"
    End If
    TBCompras_Lista!Prazo = txtPrazo_serv
    
    TBCompras_Lista!Obs_pedido = IIf(txtObs_serv = "", Null, txtObs_serv)
    TBCompras_Lista!Tipo = "S"
    TBCompras_Lista!ID_CFOP = IIf(Txt_ID_CFOP_serv = "", 0, Txt_ID_CFOP_serv)
    
    'Calcula quantidade se a unidade for diferente
    If cmbUn_serv <> Cmb_un_com_serv Then
        If FunVerifUNConversao(cmbUn_serv, Cmb_un_com_serv) = True Then
            TBCompras_Lista!Qtde_estoque = FunConverteUN(cmbUn_serv, Cmb_un_com_serv, txtQtde_serv, txtCodigo)
        Else
            TBCompras_Lista!Qtde_estoque = txtQtde_serv / FunVerificaTabelaConversaoUnidade(cmbUn_serv, Cmb_un_com_serv)
        End If
    Else
        TBCompras_Lista!Qtde_estoque = Null
    End If
End If

TBCompras_Lista.Update
txtIDLista_serv = TBCompras_Lista!IDlista

OF = IIf(txtOrdem_serv = "", 0, txtOrdem_serv)
If Cmb_OS_serv = "" Then TextoFiltro = "OS = Null" Else TextoFiltro = "OS = " & Cmb_OS_serv
Conexao.Execute "Update NFP Set NFP.Ordem = " & OF & " from tbl_Detalhes_Nota NFP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFPP.ID_carteira = " & TBCompras_Lista!IDlista & " and NFPP.Codinterno = '" & txtCodigo & "'"
Conexao.Execute "Update compras_pedido_lista Set Ordem = " & OF & ", " & TextoFiltro & " where idlista = " & txtIDLista_serv

ProcAgregarProdutoForn txtcodproduto_serv, txtIDfornecedor, txtValorUnit_serv

ProcAtualizalistaServ
If Novo_PC2 = True Then
    USMsgBox ("Novo serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo serviço"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar serviço"
    If CodigoLista3 <> 0 And ListaServ.ListItems.Count <> 0 Then
        ListaServ.SelectedItem = ListaServ.ListItems(CodigoLista3)
        ListaServ.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Pedido"
ID_documento = txtIDLista_serv
Documento = "Nº pedido: " & txtPedido
Documento1 = "Cód. interno: " & txtCodigo
ProcGravaEvento
'==================================
Novo_PC2 = False

'Recalcular centro de custo
valor = txtValorTotal_serv
NovoValor = Replace(valor, ",", ".")
Conexao.Execute "Update compras_pedido_lista_custo Set valor = (ISNULL(Percentual, 0) * " & NovoValor & ") / 100 where IDlista = " & txtIDLista_serv

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAtualizalista()
On Error GoTo tratar_erro

Listprod.ListItems.Clear
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from compras_pedido_lista where idpedido = " & txtIDPedido & " and tipo = 'P' order by idlista", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    PBLista(1).Min = 0
    PBLista(1).Max = TBCompras_Lista.RecordCount
    PBLista(1).Value = 1
    Contador = 0
    Do While TBCompras_Lista.EOF = False
        With Listprod.ListItems
            .Add , , TBCompras_Lista!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
            
            If TBCompras_Lista!Un <> TBCompras_Lista!Unidade_com Then valor = FunConversaoFinalUn(TBCompras_Lista!Un, TBCompras_Lista!Unidade_com, TBCompras_Lista!Quant_Comp, TBCompras_Lista!Desenho, True) Else valor = TBCompras_Lista!Quant_Comp
            .Item(.Count).SubItems(3) = FunFormataCasasDecimais(4, valor)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBCompras_Lista!Quant_Comp), "", FunFormataCasasDecimais(4, TBCompras_Lista!Quant_Comp))
            
            .Item(.Count).SubItems(5) = IIf(IsNull(TBCompras_Lista!preco_unitario), "", FunFormataCasasDecimais(10, TBCompras_Lista!preco_unitario))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBCompras_Lista!Desconto), "", TBCompras_Lista!Desconto)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBCompras_Lista!ValorDesconto), "0,00000", FunFormataCasasDecimais(10, TBCompras_Lista!ValorDesconto))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), "0,00000", FunFormataCasasDecimais(10, TBCompras_Lista!preco_unitario_desconto))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBCompras_Lista!IPI), "", TBCompras_Lista!IPI)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBCompras_Lista!ICMS), "", TBCompras_Lista!ICMS)
            .Item(.Count).SubItems(11) = IIf(IsNull(TBCompras_Lista!VlrIPI), "", FunFormataCasasDecimais(2, TBCompras_Lista!VlrIPI))
            .Item(.Count).SubItems(12) = IIf(IsNull(TBCompras_Lista!preco_total), "", FunFormataCasasDecimais(2, TBCompras_Lista!preco_total))
            
            If TBCompras_Lista!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBCompras_Lista!Status_Item = "APROVADO" Or TBCompras_Lista!Status_Item = "RECEBIDO" Or TBCompras_Lista!Status_Item = "CANCELADO" Then
                Status_Item = TBCompras_Lista!Status_Item
            ElseIf TBCompras_Lista!Status_Item = "N_RECEBIDO" Then
                    Status_Item = "COMPRADO"
                Else
                    Status_Item = "RECEBIDO PARCIAL"
            End If
            .Item(.Count).SubItems(13) = Status_Item
        End With
        TBCompras_Lista.MoveNext
        Contador = Contador + 1
        PBLista(1).Value = Contador
    Loop
End If
TBCompras_Lista.Close
ProcGravarTotaisPC (txtIDPedido)
ProcPuxaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Custo()
On Error GoTo tratar_erro

Lista_custo.ListItems.Clear
Qtde = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPLC.*, US.Codigo, US.Setor from Compras_pedido_lista_custo CPLC INNER JOIN Usuarios_setor US ON CPLC.ID_CC = US.ID where CPLC.IDpedido = " & txtIDPedido & " and CPLC.idlista = " & TXTIDLista & " order by US.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(1).Min = 0
    PBLista(1).Max = TBLISTA.RecordCount
    PBLista(1).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_custo.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "0,00", Format(TBLISTA!valor, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Percentual), "0,00", Format(TBLISTA!Percentual, "###,##0.0000000000"))
            .Item(.Count).SubItems(5) = TBLISTA!ID_CC
            Qtde = Qtde + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(1).Value = Contador
    Loop
End If
TBLISTA.Close
txtTotalCentro = Format(Qtde, "###,##0.00")
Qtd = IIf(txtVlrTotal_centro = "", 0, txtVlrTotal_centro)
txtSaldoCentro = Format(Qtd - Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_CustoServ()
On Error GoTo tratar_erro

Lista_custoServ.ListItems.Clear
Qtde = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPLC.*, US.Codigo, US.Setor from Compras_pedido_lista_custo CPLC INNER JOIN Usuarios_setor US ON CPLC.ID_CC = US.ID where CPLC.IDpedido = " & txtIDPedido & " and CPLC.idlista = " & txtIDLista_serv & " order by US.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(1).Min = 0
    PBLista(1).Max = TBLISTA.RecordCount
    PBLista(1).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_custoServ.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "0,00", Format(TBLISTA!valor, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Percentual), "0,00", Format(TBLISTA!Percentual, "###,##0.0000000000"))
            .Item(.Count).SubItems(5) = TBLISTA!ID_CC
            Qtde = Qtde + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(1).Value = Contador
    Loop
End If
TBLISTA.Close
txtTotalCentroServ = Format(Qtde, "###,##0.00")
Qtd = IIf(txtVlrTotal_centroServ = "", 0, txtVlrTotal_centroServ)
txtSaldoCentroServ = Format(Qtd - Qtde, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaEmpenhosProd()
On Error GoTo tratar_erro

Valor3 = 0
Lista_empenhos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VC.*, CPLE.ID, CPLE.Qtde_empenho, CPLE.Qtde_recebida FROM vendas_carteira VC INNER JOIN Compras_pedido_lista_empenhos CPLE on VC.codigo = CPLE.IDCarteira where CPLE.IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_empenhos.ListItems.Add(, , TBLISTA!ID)
            .SubItems(1) = TBLISTA!CODIGO
            
            Set TBCFOP = CreateObject("adodb.recordset")
            If IsNull(TBLISTA!ID_solicitacao) = True Or TBLISTA!ID_solicitacao = 0 Then
                TBCFOP.Open "Select Ncotacao, Revisao, cliente FROM vendas_proposta where cotacao = " & TBLISTA!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    .SubItems(2) = IIf(IsNull(TBCFOP!Ncotacao), "", TBCFOP!Ncotacao)
                    .SubItems(3) = IIf(IsNull(TBCFOP!Revisao), "", TBCFOP!Revisao)
                    .SubItems(4) = IIf(IsNull(TBCFOP!Cliente), "", TBCFOP!Cliente)
                End If
            Else
                TBCFOP.Open "Select Requisicaotexto FROM Outros_SolicitacaoPCP where ID = " & TBLISTA!ID_solicitacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then .SubItems(2) = IIf(IsNull(TBCFOP!Requisicaotexto), "", TBCFOP!Requisicaotexto)
            End If
            TBCFOP.Close
            
            .SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .SubItems(6) = IIf(IsNull(TBLISTA!Rev_codinterno), "", TBLISTA!Rev_codinterno)
            .SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .SubItems(8) = IIf(IsNull(TBLISTA!descricao_tecnica), "", TBLISTA!descricao_tecnica)
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .SubItems(9) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_recebida), 0, TBLISTA!Qtde_recebida)
            .SubItems(10) = Format(Valor1, "###,##0.0000")
            .SubItems(11) = Format(valor - Valor1, "###,##0.0000")
            .SubItems(12) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .SubItems(13) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .SubItems(14) = IIf(IsNull(TBLISTA!PCCliente), "", TBLISTA!PCCliente)
            .SubItems(15) = IIf(IsNull(TBLISTA!N_item), "", TBLISTA!N_item)
            .SubItems(16) = IIf(IsNull(TBLISTA!Liberacao), "", TBLISTA!Liberacao)
        End With
        Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close
valor = Txt_qtde_total_comprada(0)
Txt_qtde_total_emp(0) = Format(Valor3, "###,##0.0000")
Txt_qtde_total_disp(0) = Format(valor - Valor3, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaEmpenhosServ()
On Error GoTo tratar_erro

Valor3 = 0
Lista_empenhos_serv.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VC.*, CPLE.ID, CPLE.Qtde_empenho, CPLE.Qtde_recebida FROM vendas_carteira VC INNER JOIN Compras_pedido_lista_empenhos CPLE on VC.codigo = CPLE.IDCarteira where CPLE.IDlista = " & txtIDLista_serv, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_empenhos_serv.ListItems.Add(, , TBLISTA!ID)
            .SubItems(1) = TBLISTA!CODIGO
            
            Set TBCFOP = CreateObject("adodb.recordset")
            If IsNull(TBLISTA!ID_solicitacao) = True Or TBLISTA!ID_solicitacao = 0 Then
                TBCFOP.Open "Select Ncotacao, Revisao, cliente FROM vendas_proposta where cotacao = " & TBLISTA!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    .SubItems(2) = IIf(IsNull(TBCFOP!Ncotacao), "", TBCFOP!Ncotacao)
                    .SubItems(3) = IIf(IsNull(TBCFOP!Revisao), "", TBCFOP!Revisao)
                    .SubItems(4) = IIf(IsNull(TBCFOP!Cliente), "", TBCFOP!Cliente)
                End If
            Else
                TBCFOP.Open "Select Requisicaotexto FROM Outros_SolicitacaoPCP where ID = " & TBLISTA!ID_solicitacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then .SubItems(2) = IIf(IsNull(TBCFOP!Requisicaotexto), "", TBCFOP!Requisicaotexto)
            End If
            TBCFOP.Close
            
            .SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .SubItems(6) = IIf(IsNull(TBLISTA!Rev_codinterno), "", TBLISTA!Rev_codinterno)
            .SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .SubItems(8) = IIf(IsNull(TBLISTA!descricao_tecnica), "", TBLISTA!descricao_tecnica)
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .SubItems(9) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_recebida), 0, TBLISTA!Qtde_recebida)
            .SubItems(10) = Format(Valor1, "###,##0.0000")
            .SubItems(11) = Format(valor - Valor1, "###,##0.0000")
            .SubItems(12) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .SubItems(13) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .SubItems(14) = IIf(IsNull(TBLISTA!PCCliente), "", TBLISTA!PCCliente)
            .SubItems(15) = IIf(IsNull(TBLISTA!N_item), "", TBLISTA!N_item)
            .SubItems(16) = IIf(IsNull(TBLISTA!Liberacao), "", TBLISTA!Liberacao)
        End With
        Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close
valor = Txt_qtde_total_comprada(1)
Txt_qtde_total_emp(1) = Format(Valor3, "###,##0.0000")
Txt_qtde_total_disp(1) = Format(valor - Valor3, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAtualizalistaServ()
On Error GoTo tratar_erro

ListaServ.ListItems.Clear
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from compras_pedido_lista where idpedido = " & txtIDPedido & " and tipo = 'S' order by idlista", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    TBCompras_Lista.MoveLast
    PBLista(1).Min = 0
    PBLista(1).Max = TBCompras_Lista.RecordCount
    PBLista(1).Value = 1
    Contador = 0
    TBCompras_Lista.MoveFirst
    Do While TBCompras_Lista.EOF = False
        With ListaServ.ListItems
            .Add , , TBCompras_Lista!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBCompras_Lista!Desenho), "", (Trim(TBCompras_Lista!Desenho)))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBCompras_Lista!Descricao), "", (Trim(TBCompras_Lista!Descricao)))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBCompras_Lista!Quant_Comp), "", (Format(TBCompras_Lista!Quant_Comp, "0.000")))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBCompras_Lista!preco_unitario), "", (Format(TBCompras_Lista!preco_unitario, "###,##0.0000000000")))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBCompras_Lista!Desconto), 0, TBCompras_Lista!Desconto)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBCompras_Lista!ValorDesconto), "0,00000", (Format(TBCompras_Lista!ValorDesconto, "###,##0.0000000000")))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), "0,00000", (Format(TBCompras_Lista!preco_unitario_desconto, "###,##0.0000000000")))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBCompras_Lista!ISSQN), "", TBCompras_Lista!ISSQN)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBCompras_Lista!VlrISSQN), "", (Format(TBCompras_Lista!VlrISSQN, "###,##0.00")))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBCompras_Lista!preco_total), "", (Format(TBCompras_Lista!preco_total, "###,##0.00")))
            
            If TBCompras_Lista!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBCompras_Lista!Status_Item = "APROVADO" Or TBCompras_Lista!Status_Item = "RECEBIDO" Or TBCompras_Lista!Status_Item = "CANCELADO" Then
                Status_Item = TBCompras_Lista!Status_Item
            ElseIf TBCompras_Lista!Status_Item = "N_RECEBIDO" Then
                    Status_Item = "COMPRADO"
                Else
                    Status_Item = "RECEBIDO PARCIAL"
            End If
            .Item(.Count).SubItems(11) = Status_Item
        End With
        TBCompras_Lista.MoveNext
        Contador = Contador + 1
        PBLista(1).Value = Contador
    Loop
End If
TBCompras_Lista.Close
ProcGravarTotaisPC (txtIDPedido)
ProcPuxaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEscolher_item_Click()
On Error GoTo tratar_erro

If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "produtos", "localizar", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "produtos", "localizar", True, True) = False Then Exit Sub
End If
If FunVerifSatus("localizar produtos", True) = False Then Exit Sub
ProcLiberaTabsProd
Sit_Nota = 1
frmCompras_ListaProduto.Show 1
If txtEspecificacoes <> "" Then txtprazo_item.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirTab_Produto()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0: ProcExcluirLista
    Case 1: ProcExcluirLista_Custo
    Case 2: ProcExcluirEmpenhoProd
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
LiberarData = False

ProcCarregaToolBar1 Me, 15105, 5, True
ProcCarregaToolBar2 Me, 15195, 17, True
ProcCarregaToolBar3 Me, 15195, 9, True
ProcCarregaToolBar4 Me, 15195, 14, True
ProcCarregaToolBar5 Me, 15195, 14, True
ProcCarregaToolBar6 Me, 15195, 10, True
Cmb_opcao_lista = "Validação"
Formulario = "Compras/Pedido"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
SSTab2.Tab = 0
SSTab3.Tab = 0
SSTab4.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboEmpresa Cmb_empresa_carteira, False
ProcCarregaCombo
ProcCarregaCombos
ProcCarregaCombosServ
ProcCarregaCamposCombo

ProcFiltroPadrao cmbfiltrarpor_necess, Optmeio_necess, Optfim_necess, optIgual_necess, Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex), "Produtos/Serviços", "C", True
If Permitido = False Then cmbfiltrarpor_necess = "Código interno"
ProcFiltroPadrao cmbfiltrarpor_sol, optMeio_sol, optFim_sol, optIgual_sol, Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex), "Produtos/Serviços", "C", True
If Permitido = False Then cmbfiltrarpor_sol = "Código interno"
Cmb_filtrar = "Com necessidade"

ProcRemoveObjetosResize Me

'=======================================================
StrSql = "update compras_pedido set CPF_CNPJ = CF.CPF_CNPJ From Compras_pedido CP Inner join Compras_fornecedores CF on CP.idfornecedor = CF.IDCliente Where CP.CPF_CNPJ IS NULL"
Conexao.Execute StrSql
StrSql = ""
'=======================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombo()
On Error GoTo tratar_erro

With cmbMoeda
    .Clear
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select Moeda from moeda", Conexao, adOpenKeyset, adLockOptimistic
    If TBFamilia.EOF = False Then
        .AddItem ""
        Do While TBFamilia.EOF = False
            .AddItem TBFamilia!Moeda
            TBFamilia.MoveNext
        Loop
    End If
    TBFamilia.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombos()
On Error GoTo tratar_erro

ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombosServ()
On Error GoTo tratar_erro

ProcCarregaComboUnidade cmbUn_serv, False
ProcCarregaComboUnidade Cmb_un_com_serv, False
ProcCarregaComboFamilia cmbFamilia_serv, "familia <> 'Null' and Compras = 'True'", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

If IsNull(TBCompras_Pedido!ID_empresa) = False And TBCompras_Pedido!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBCompras_Pedido!ID_empresa
Caption = "Compras - Pedido - (Pedido : " & IIf(IsNull(TBCompras_Pedido!Pedido), "", TBCompras_Pedido!Pedido) & ")"
txtIDPedido = TBCompras_Pedido!IDpedido

With txtPedido
    .Text = TBCompras_Pedido!Pedido
    .Locked = True
    .TabStop = False
End With

Select Case TBCompras_Pedido!Status_pedido
    Case "ABERTO": txtStatus.Text = "COMPRADO"
    Case "PARCIAL": txtStatus.Text = "RECEBIDO PARCIAL"
    Case "ENCERRADO": txtStatus.Text = "RECEBIDO"
    Case "CANCELADO": txtStatus.Text = "CANCELADO"
    Case "AGUARDANDO APROVAÇÃO": txtStatus.Text = "AGUARDANDO APROVAÇÃO"
    Case "APROVADO": txtStatus.Text = "APROVADO"
End Select
If TBCompras_Pedido!Email_Enviado = True Then Chk_email_enviado.Value = 1 Else Chk_email_enviado.Value = 0
txtData.Text = IIf(IsNull(TBCompras_Pedido!Data), "", Format(TBCompras_Pedido!Data, "dd/mm/yy"))
txtResponsavel.Text = IIf(IsNull(TBCompras_Pedido!Responsavel), "", TBCompras_Pedido!Responsavel)
txtDtValidacao = IIf(IsNull(TBCompras_Pedido!DtValidacao), "", TBCompras_Pedido!DtValidacao)
txtRespValidacao = IIf(IsNull(TBCompras_Pedido!RespValidacao), "", TBCompras_Pedido!RespValidacao)
txtData_aprovacao.Text = IIf(IsNull(TBCompras_Pedido!Data_aprovado), "", TBCompras_Pedido!Data_aprovado)
txtResponsavel_aprovacao.Text = IIf(IsNull(TBCompras_Pedido!Resp_aprovado), "", TBCompras_Pedido!Resp_aprovado)
txtIDfornecedor.Text = IIf(IsNull(TBCompras_Pedido!IDFornecedor), "", TBCompras_Pedido!IDFornecedor)

txtcnpj.Text = IIf(IsNull(TBCompras_Pedido!CPF_CNPJ), "", TBCompras_Pedido!CPF_CNPJ)

txtContato = IIf(IsNull(TBCompras_Pedido!contato), "", TBCompras_Pedido!contato)
txtTipo_endereco = IIf(IsNull(TBCompras_Pedido!Tipo_endereco), "", TBCompras_Pedido!Tipo_endereco)
txtendereco = IIf(IsNull(TBCompras_Pedido!Endereco), "", TBCompras_Pedido!Endereco)
txtNumero = IIf(IsNull(TBCompras_Pedido!Numero), "", TBCompras_Pedido!Numero)
txtBairro = IIf(IsNull(TBCompras_Pedido!Bairro), "", TBCompras_Pedido!Bairro)
txtTipo_bairro = IIf(IsNull(TBCompras_Pedido!Tipo_bairro), "", TBCompras_Pedido!Tipo_bairro)
txtCidade = IIf(IsNull(TBCompras_Pedido!Cidade), "", TBCompras_Pedido!Cidade)
txtuf = IIf(IsNull(TBCompras_Pedido!Estado), "", TBCompras_Pedido!Estado)
txtEmail = IIf(IsNull(TBCompras_Pedido!Email), "", TBCompras_Pedido!Email)
txttelefone = IIf(IsNull(TBCompras_Pedido!fone), "", TBCompras_Pedido!fone)
txtFax = IIf(IsNull(TBCompras_Pedido!Fax), "", TBCompras_Pedido!Fax)
Txt_n_referencia = IIf(IsNull(TBCompras_Pedido!N_referencia), "", TBCompras_Pedido!N_referencia)
Txt_descricao_referencia = IIf(IsNull(TBCompras_Pedido!Descricao_referencia), "", TBCompras_Pedido!Descricao_referencia)

Novo_PC = False
Frame1(16).Enabled = True
Frame1(4).Enabled = True
ProcLimparTudo

ProcPuxaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosLista()
On Error GoTo tratar_erro

txtNomenclatura.Text = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)

If IsNull(TBProduto!ID_Requisicao) = False And TBProduto!ID_Requisicao <> 0 Then
    ProcCarregaComboCodRef cmbReferencia, "P.desenho = '" & TBProduto!Desenho & "'", 0, "", False, True
Else
    ProcCarregaComboCodRef cmbReferencia, "P.desenho = '" & TBProduto!Desenho & "'", txtIDfornecedor, "F", True, True
End If

NomeCampo = "a unidade de estoque"
If IsNull(TBProduto!Un) = False And TBProduto!Un <> "" Then cmbun.Text = TBProduto!Un
NomeCampo = "a unidade comercial"
If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com.Text = TBProduto!Unidade_com
NomeCampo = "a família"
If IsNull(TBProduto!Familia) = False And TBProduto!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbfamilia.Text = TBProduto!Familia
    VerifDadosPadraoFamilia = True
End If
NomeCampo = "o código de referência"
If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" Then cmbReferencia = TBProduto!N_referencia Else cmbReferencia.ListIndex = -1
txtOrdem = IIf(IsNull(TBProduto!Ordem), "", TBProduto!Ordem)

If FunVerifOPCarregaOS(Cmb_OS, txtOrdem, False, True) = True Then
    NomeCampo = "a OS"
    If IsNull(TBProduto!OS) = False And TBProduto!OS <> "" Then Cmb_OS = TBProduto!OS
End If

ProcCarregaDadosCFOPProdServ IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), True
ProcCarregaCST
NomeCampo = "a CST de ICMS"
If IsNull(TBProduto!CST) = False And TBProduto!CST <> "" Then Cmb_CST_ICMS.Text = TBProduto!CST

1:
    If TBProduto!Desconto > 0 Then Chk_desc.Value = 1
    
    TXTIDLista.Text = Listprod.SelectedItem
    Txt_vlr_unit_ultima_compra_prod = FunVerifVlrUnitUltCompra(txtNomenclatura, TXTIDLista)
    txtcodproduto.Text = IIf(IsNull(TBProduto!Codproduto), "0", TBProduto!Codproduto)
    If IsNull(TBProduto!Status_Item) = False And TBProduto!Status_Item <> "" Then
        txtstatus_item = Listprod.SelectedItem.ListSubItems(13)
'        If TBProduto!Status_Item = "RECEBIDO" Or TBProduto!Status_Item = "PARCIAL" Or TBProduto!Status_Item = "CANCELADO" Or IsNull(TBProduto!ID_Requisicao) = False And TBProduto!ID_Requisicao <> 0 Or IsNull(TBProduto!ID_cotacao) = False And TBProduto!ID_cotacao <> 0 Then
'            If TBProduto!Status_Item = "RECEBIDO" Or TBProduto!Status_Item = "PARCIAL" Then ProcBloqueiaCamposItem Else ProcDesbloqueiaCamposItem
'            If IsNull(TBProduto!ID_Requisicao) = False And TBProduto!ID_Requisicao <> 0 Or IsNull(TBProduto!ID_cotacao) = False And TBProduto!ID_cotacao <> 0 Then ProcBloqueiaCamposItemSolCot
'            ProcBloqLibQtde
'            ProcBloqueiaTabsProd
'        Else
'            If TBProduto!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBProduto!Status_Item = "N_RECEBIDO" Then ProcDesbloqueiaCamposItem Else ProcBloqueiaCamposItem
'            If IsNull(TBProduto!ID_Requisicao) = True Or TBProduto!ID_Requisicao = 0 Or IsNull(TBProduto!ID_cotacao) = True Or TBProduto!ID_cotacao = 0 Then ProcDesbloqueiaCamposItemSolCot
'            ProcBloqLibQtde
'            ProcLiberaTabsProd
'        End If
    End If
    txtQuantidade.Text = IIf(IsNull(TBProduto!Quant_Comp), "", Format(TBProduto!Quant_Comp, "###,##0.0000"))
    txtQuantidade_PC = IIf(IsNull(TBProduto!Quant_Comp_PC), "", TBProduto!Quant_Comp_PC)
    txtEspecificacoes.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    txtDescricao_comercial.Text = IIf(IsNull(TBProduto!Descricao_comercial), "", TBProduto!Descricao_comercial)
    txtdetalheitem.Text = IIf(IsNull(TBProduto!detalheitem), "", TBProduto!detalheitem)
    txtObs.Text = IIf(IsNull(TBProduto!Obs_pedido), "", TBProduto!Obs_pedido)
    txtvalorunitario.Text = Format(TBProduto!preco_unitario, "###,##0.0000000000")
    txtDesconto = IIf(IsNull(TBProduto!Desconto), "", TBProduto!Desconto)
    txtvalordesconto = IIf(IsNull(TBProduto!ValorDesconto), "", Format(TBProduto!ValorDesconto, "###,##0.0000000000"))
    If IsNull(TBProduto!preco_unitario_desconto) = True Then
        txtvalorunitariodesc.Text = txtvalorunitario.Text
    Else
        If TBProduto!preco_unitario_desconto = 0 Then
            txtvalorunitariodesc.Text = txtvalorunitario.Text
        Else
            txtvalorunitariodesc.Text = Format(TBProduto!preco_unitario_desconto, "###,##0.0000000000")
        End If
    End If
    txtprazo_item = IIf(IsNull(TBProduto!Prazo), "__/__/____", Format(TBProduto!Prazo, "dd/mm/yyyy"))
    If TBProduto!Remessa = True Then chkRemessa.Value = 1 Else chkRemessa.Value = 0
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from projproduto where desenho = '" & TBProduto!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        With txtvalorunitario
            If TBAbrir!Valor_bloqueado = True Or TBProduto!Status_Item = "RECEBIDO" Or TBProduto!Status_Item = "PARCIAL" Or TBProduto!Status_Item = "CANCELADO" Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
        ProcBloqueiaCamposProdComCadastrado
    Else
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposProdSemCadastrado
    End If
    TBAbrir.Close

    Txt_ID_CF = IIf(IsNull(TBProduto!ID_CF), 0, TBProduto!ID_CF)
    
'    If Txt_ID_CF <> "" And Txt_ID_CF <> "0" Then
'        ProcValorImposto txtpedido, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtfornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(txt_ID_CFOP_prod = "", 0, txt_ID_CFOP_prod), 0
'
'        ProcControleImposto IIf(txt_ID_CFOP_prod = "", 0, txt_ID_CFOP_prod), 0
'        If TemIPI = "SIM" Then txtipi = IntIPI Else txtipi.Text = 0
'        If TemICMS = "SIM" Then txtIcms = IntICMS Else txtIcms.Text = 0
'    Else
'        txtipi.Text = IIf(IsNull(TBProduto!IPI), 0, TBProduto!IPI)
'        TxtvlrIpi.Text = IIf(IsNull(TBProduto!VlrIPI), "", Format(TBProduto!VlrIPI, "###,##0.00"))
'        txtIcms = IIf(IsNull(TBProduto!ICMS), "", TBProduto!ICMS)
'        TxtVlrIcms = IIf(IsNull(TBProduto!vlrICMS), "", Format(TBProduto!vlrICMS, "###,##0.00"))
'    End If
    
    txtIPI.Text = IIf(IsNull(Listprod.SelectedItem.ListSubItems.Item(9).Text), 0, Listprod.SelectedItem.ListSubItems.Item(9).Text)
    TxtvlrIpi.Text = IIf(IsNull(Listprod.SelectedItem.ListSubItems.Item(11).Text), 0, Listprod.SelectedItem.ListSubItems.Item(11).Text)
    txtICMS.Text = IIf(IsNull(Listprod.SelectedItem.ListSubItems.Item(10).Text), 0, Listprod.SelectedItem.ListSubItems.Item(10).Text)
    txtvlrICMS.Text = IIf(IsNull(Listprod.SelectedItem.ListSubItems.Item(12).Text), 0, Listprod.SelectedItem.ListSubItems.Item(12).Text)
    
    txtFrete = IIf(IsNull(TBProduto!Frete), "", Format(TBProduto!Frete, "###,##0.00"))
    txtSeguro = IIf(IsNull(TBProduto!Seguro), "", Format(TBProduto!Seguro, "###,##0.00"))
    txtAcessorias = IIf(IsNull(TBProduto!Acessorias), "", Format(TBProduto!Acessorias, "###,##0.00"))
    If TBProduto!Frete_IPI = True Then ChkFrete_IPI.Value = 1 Else ChkFrete_IPI.Value = 0
    
    'Centro de custo
    ProcLimpaCamposCusto
    Frame1(13).Enabled = False
   
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        If NomeCampo = "a OS" Then NomeCampo1 = TBProduto!OS Else NomeCampo1 = ""
        USMsgBox ("Não foi encontrado " & NomeCampo & " " & NomeCampo1 & " deste produto."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosListaServ()
On Error GoTo tratar_erro

txtCodigo.Text = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)

ProcCarregaComboCodRef cmbreferencia_serv, "P.desenho = '" & TBProduto!Desenho & "'", txtIDfornecedor, "F", True, True
NomeCampo = "a unidade de estoque"
If IsNull(TBProduto!Un) = False And TBProduto!Un <> "" Then cmbUn_serv.Text = TBProduto!Un
NomeCampo = "a unidade comercial"
If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com_serv.Text = TBProduto!Unidade_com
NomeCampo = "a família"
If IsNull(TBProduto!Familia) = False And TBProduto!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbFamilia_serv.Text = TBProduto!Familia
    VerifDadosPadraoFamilia = True
End If
NomeCampo = "o código de referência"
If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" Then cmbreferencia_serv = TBProduto!N_referencia Else cmbreferencia_serv.ListIndex = -1
txtOrdem_serv = IIf(IsNull(TBProduto!Ordem), "", TBProduto!Ordem)

If FunVerifOPCarregaOS(Cmb_OS_serv, txtOrdem_serv, False, True) = True Then
    NomeCampo = "a OS"
    If IsNull(TBProduto!OS) = False And TBProduto!OS <> "" Then Cmb_OS_serv = TBProduto!OS
End If

1:
    If TBProduto!Desconto > 0 Then Chk_desc2.Value = 1
    
    txtIDLista_serv.Text = ListaServ.SelectedItem
    Txt_vlr_unit_ultima_compra_serv = FunVerifVlrUnitUltCompra(txtCodigo, txtIDLista_serv)
    txtcodproduto_serv = IIf(IsNull(TBProduto!Codproduto), "0", TBProduto!Codproduto)
    If IsNull(TBProduto!Status_Item) = False And TBProduto!Status_Item <> "" Then
        txtStatus_serv = ListaServ.SelectedItem.ListSubItems(11)
        If TBProduto!Status_Item = "RECEBIDO" Or TBProduto!Status_Item = "PARCIAL" Or TBProduto!Status_Item = "CANCELADO" Or IsNull(TBProduto!ID_Requisicao) = False And TBProduto!ID_Requisicao <> 0 Or IsNull(TBProduto!ID_cotacao) = False And TBProduto!ID_cotacao <> 0 Then
            If TBProduto!Status_Item = "RECEBIDO" Or TBProduto!Status_Item = "PARCIAL" Then ProcBloqueiaCamposServ Else ProcDesbloqueiaCamposServ
            If IsNull(TBProduto!ID_Requisicao) = False And TBProduto!ID_Requisicao <> 0 Or IsNull(TBProduto!ID_cotacao) = False And TBProduto!ID_cotacao <> 0 Then ProcBloqueiaCamposServSolCot
            ProcBloqueiaTabsServ
        Else
            If Permitido = True Then
                If TBProduto!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBProduto!Status_Item = "N_RECEBIDO" Then ProcDesbloqueiaCamposServ Else ProcBloqueiaCamposServ
                If IsNull(TBProduto!ID_Requisicao) = True Or TBProduto!ID_Requisicao = 0 Or IsNull(TBProduto!ID_cotacao) = True Or TBProduto!ID_cotacao = 0 Then ProcDesbloqueiaCamposServSolCot
                ProcLiberaTabsServ
            End If
        End If
    End If
        
    txtQtde_serv.Text = IIf(IsNull(TBProduto!Quant_Comp), "", Format(TBProduto!Quant_Comp, "###,##0.0000"))
    txtDescricao_serv.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    txtDescricao_comercialServ.Text = IIf(IsNull(TBProduto!Descricao_comercial), "", TBProduto!Descricao_comercial)
    txtDetalhe_serv.Text = IIf(IsNull(TBProduto!detalheitem), "", TBProduto!detalheitem)
    txtObs_serv.Text = IIf(IsNull(TBProduto!Obs_pedido), "", TBProduto!Obs_pedido)
    txtValorUnit_serv.Text = Format(TBProduto!preco_unitario, "###,##0.0000000000")
    txtDesconto_serv = IIf(IsNull(TBProduto!Desconto), "", TBProduto!Desconto)
    txtVlrDesconto_serv = IIf(IsNull(TBProduto!ValorDesconto), "", Format(TBProduto!ValorDesconto, "###,##0.0000000000"))
    If IsNull(TBProduto!preco_unitario_desconto) = True Then
        txtVlrUnitDesc_serv.Text = txtValorUnit_serv.Text
    Else
        If TBProduto!preco_unitario_desconto = 0 Then
            txtVlrUnitDesc_serv.Text = txtValorUnit_serv.Text
        Else
            txtVlrUnitDesc_serv.Text = Format(TBProduto!preco_unitario_desconto, "###,##0.0000000000")
        End If
    End If
    txtISSQN = IIf(IsNull(TBProduto!ISSQN), "", TBProduto!ISSQN)
    txtValor_ISSQN = IIf(IsNull(TBProduto!VlrISSQN), "", Format(TBProduto!VlrISSQN, "###,##0.00"))
    txtPrazo_serv.Text = IIf(IsNull(TBProduto!Prazo), "__/__/____", Format(TBProduto!Prazo, "dd/mm/yyyy"))
    ProcCarregaDadosCFOPProdServ IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), False
        
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from projproduto where desenho = '" & TBProduto!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        With txtValorUnit_serv
            If TBAbrir!Valor_bloqueado = True Or TBProduto!Status_Item = "RECEBIDO" Or TBProduto!Status_Item = "PARCIAL" Or TBProduto!Status_Item = "CANCELADO" Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
        ProcBloqueiaCamposServComCadastrado
    Else
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposServSemCadastrado
    End If
    TBAbrir.Close
    
    'Centro de custo
    ProcLimpaCamposCustoServ
    Frame1(8).Enabled = False
    
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste serviço."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBCompras_Pedido!Pedido = txtPedido.Text
TBCompras_Pedido!ID_empresa = IIf(SSTab1.Tab = 0, Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex), Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
TBCompras_Pedido!IDFornecedor = txtIDfornecedor.Text
TBCompras_Pedido!Fornecedor = txtFornecedor.Text
TBCompras_Pedido!CPF_CNPJ = IIf(txtcnpj.Text = "", Null, txtcnpj.Text)
TBCompras_Pedido!Categoria = IIf(txtCategoria = "", Null, txtCategoria)
TBCompras_Pedido!contato = IIf(txtContato = "", Null, txtContato)
TBCompras_Pedido!Tipo_endereco = IIf(txtTipo_endereco = "", Null, txtTipo_endereco)
TBCompras_Pedido!Endereco = IIf(txtendereco = "", Null, txtendereco)
TBCompras_Pedido!Numero = IIf(txtNumero = "", Null, txtNumero)
TBCompras_Pedido!Tipo_bairro = IIf(txtTipo_bairro = "", Null, txtTipo_bairro)
TBCompras_Pedido!Bairro = IIf(txtBairro = "", Null, txtBairro)
TBCompras_Pedido!Cidade = IIf(txtCidade = "", Null, txtCidade)
TBCompras_Pedido!Estado = IIf(txtuf = "", Null, txtuf)
TBCompras_Pedido!Email = IIf(txtEmail.Text = "", Null, txtEmail.Text)
TBCompras_Pedido!fone = IIf(txttelefone = "", Null, txttelefone)
TBCompras_Pedido!Fax = IIf(txtFax = "", Null, txtFax)
TBCompras_Pedido!N_referencia = IIf(Txt_n_referencia = "", Null, Txt_n_referencia)
TBCompras_Pedido!Descricao_referencia = IIf(Txt_descricao_referencia = "", Null, Txt_descricao_referencia)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro
  
frmCompras_Pedido_Menu_Impressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosComercial()
On Error GoTo tratar_erro

TBProduto!condicoes = IIf(cmbpagamento = "", Null, cmbpagamento)
TBProduto!IDpedido = txtIDPedido
TBProduto!Embalagem = IIf(txtembalagem = "", Null, txtembalagem)
TBProduto!ID_entrega = Txt_ID_entrega
TBProduto!localentrega = IIf(txtlocal = "", Null, txtlocal)
If cmbtransporte <> "" Then
    Select Case Cmb_tipo_transp
        Case "Cliente": TBProduto!Tipo_transp = "C"
        Case "Fornecedor": TBProduto!Tipo_transp = "F"
        Case "Empresa": TBProduto!Tipo_transp = "E"
    End Select
    TBProduto!Idtransporte = cmbtransporte.ItemData(cmbtransporte.ListIndex)
Else
    TBProduto!Tipo_transp = ""
    TBProduto!Idtransporte = 0
End If
TBProduto!Observacoes = IIf(txtObservacoes = "", Null, txtObservacoes)
TBProduto!Prazo = IIf(txtprazo = "", Null, txtprazo)
TBProduto!Banco = IIf(txtBanco = "", Null, txtBanco)
TBProduto!Agencia = IIf(txtAgencia = "", Null, txtAgencia)
TBProduto!Conta = IIf(txtConta = "", Null, txtConta)
TBProduto!transporte = IIf(txttransporte = "", Null, txttransporte)
TBProduto!Moeda = IIf(cmbMoeda = "", Null, cmbMoeda)
TBProduto!Valor_moeda = IIf(Txt_valor_moeda = "", Null, Txt_valor_moeda)
If chkObs_Financeiro.Value = 1 Then TBProduto!Obs_financeiro = True Else TBProduto!Obs_financeiro = False

'Atualiza valor dos produtos/serviços de acordo com valor da moeda
If IsNull(TBProduto!Valor_moeda) = False Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * FROM Compras_pedido_lista where IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            TBLISTA!preco_unitario = TBLISTA!preco_unitario / TBProduto!Valor_moeda
            TBLISTA!preco_unitario_desconto = TBLISTA!preco_unitario_desconto / TBProduto!Valor_moeda
            TBLISTA!preco_total = TBLISTA!preco_total / TBProduto!Valor_moeda
            TBLISTA.Update
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
    ProcAtualizalista
    ProcAtualizalistaServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarComercial()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "os dados comerciais", "salvar", True, True) = False Then Exit Sub
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "os dados comerciais", "salvar", True, True) = False Then Exit Sub
End If
If FunVerifSatus("salvar os dados comerciais", True) = False Then Exit Sub
Acao = "salvar"
If cmbMoeda = "" Then
    NomeCampo = "a moeda"
    ProcVerificaAcao
    cmbMoeda.SetFocus
    Exit Sub
End If
If cmbMoeda.Text = "REAL" Then Txt_valor_moeda.Text = "1,00"

valor = IIf(Txt_valor_moeda = "", 0, Txt_valor_moeda)
If valor <= 0 Then
    If cmbMoeda.Text <> "REAL" Then
    NomeCampo = "Valor do " & cmbMoeda.Text & " hoje "
    End If
    ProcVerificaAcao
    Txt_valor_moeda.SetFocus
    Exit Sub
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM compras_comercial WHERE IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dados Comerciais"
    If IsNull(TBProduto!Valor_moeda) = False Then
        If Format(valor, "###,##0.00") <> Format(TBProduto!Valor_moeda, "###,##0.00") Then
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select * FROM Compras_pedido_lista where IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    TBLISTA!preco_unitario = TBLISTA!preco_unitario / TBProduto!Valor_moeda
                    TBLISTA!preco_unitario_desconto = TBLISTA!preco_unitario_desconto / TBProduto!Valor_moeda
                    TBLISTA!preco_total = TBLISTA!preco_total / TBProduto!Valor_moeda
                    TBLISTA.Update
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
            If valor = 0 Then
                ProcAtualizalista
                ProcAtualizalistaServ
            End If
        End If
    End If
Else
    TBProduto.AddNew
    USMsgBox ("Dados comerciais salvos com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo dados Comerciais"
End If
ProcEnviadadosComercial
TBProduto.Update
TBProduto.Close
'==================================
Modulo = "Compras/Pedido"
ID_documento = txtIDPedido
Documento = "Nº pedido: " & txtPedido.Text
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAbreComercial()
On Error GoTo tratar_erro
Dim PrecoTotal As Double 'OK

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * FROM compras_comercial WHERE idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    ProcPuxaDadosComercial
End If
TBCompras.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparComercial()
On Error GoTo tratar_erro

cmbpagamento = "N/A"
txtprazo.Text = "N/A"
Txt_ID_entrega = 0
txtlocal.Text = "N/A"
Cmb_tipo_transp.ListIndex = -1
cmbtransporte.ListIndex = -1
txtembalagem.Text = "N/A"
txtObservacoes.Text = ""
chkObs_Financeiro.Value = 0
txtBanco = ""
txtAgencia = ""
txtConta = ""
txttransporte.Text = "N/A"
cmbMoeda.ListIndex = -1
Txt_valor_moeda = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosComercial()
On Error GoTo tratar_erro

cmbpagamento.Text = IIf(IsNull(TBCompras!condicoes), "", TBCompras!condicoes)
txtembalagem.Text = IIf(IsNull(TBCompras!Embalagem), "", TBCompras!Embalagem)
txtObservacoes.Text = IIf(IsNull(TBCompras!Observacoes), "", TBCompras!Observacoes)
txtprazo.Text = IIf(IsNull(TBCompras!Prazo), "", TBCompras!Prazo)
Txt_ID_entrega = IIf(IsNull(TBCompras!ID_entrega), 0, TBCompras!ID_entrega)
txtlocal.Text = IIf(IsNull(TBCompras!localentrega), "", TBCompras!localentrega)
If TBCompras!Obs_financeiro = True Then chkObs_Financeiro.Value = 1 Else chkObs_Financeiro.Value = 0
If IsNull(TBCompras!Idtransporte) = False And TBCompras!Idtransporte <> "0" Then
    If TBCompras!Tipo_transp = "E" Then
        Cmb_tipo_transp = "Empresa"
        NomeTabela = "Empresa"
        NomeCampo = "Codigo"
    Else
        If TBCompras!Tipo_transp = "C" Then
            Cmb_tipo_transp = "Cliente"
            NomeTabela = "Clientes"
        Else
            Cmb_tipo_transp = "Fornecedor"
            NomeTabela = "Compras_fornecedores"
        End If
        NomeCampo = "IDcliente"
    End If
    Set TBTransporte = CreateObject("adodb.recordset")
    TBTransporte.Open "Select * from " & NomeTabela & " where " & NomeCampo & " = " & TBCompras!Idtransporte, Conexao, adOpenKeyset, adLockOptimistic
    If TBTransporte.EOF = False Then
        With cmbtransporte
            Select Case TBCompras!Tipo_transp
                Case "C": .Text = TBTransporte!NomeRazao
                Case "F": .Text = TBTransporte!Nome_Razao
                Case "E": .Text = TBTransporte!Empresa
            End Select
        End With
    End If
    TBTransporte.Close
End If
txtBanco = IIf(IsNull(TBCompras!Banco), "", TBCompras!Banco)
txtAgencia = IIf(IsNull(TBCompras!Agencia), "", TBCompras!Agencia)
txtConta = IIf(IsNull(TBCompras!Conta), "", TBCompras!Conta)
txttransporte = IIf(IsNull(TBCompras!transporte), "", TBCompras!transporte)
If IsNull(TBCompras!Moeda) = False And TBCompras!Moeda <> "" Then cmbMoeda = TBCompras!Moeda
Txt_valor_moeda = IIf(IsNull(TBCompras!Valor_moeda), "", Format(TBCompras!Valor_moeda, "###,##0.0000"))
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Pedido"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaCombos
ProcCarregaCombosServ
ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbtransporte.ListIndex = -1
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
cmbFamilia_serv.ListIndex = -1
cmbUn_serv.ListIndex = -1
Cmb_un_com_serv.ListIndex = -1
If txtIDPedido <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * FROM compras_comercial WHERE idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Idtransporte) = False And TBAbrir!Idtransporte <> "0" Then
            If TBAbrir!Tipo_transp = "E" Then
                NomeTabela = "Empresa"
                NomeCampo = "Codigo"
            Else
                NomeCampo = "IDcliente"
                If TBAbrir!Tipo_transp = "C" Then NomeTabela = "Clientes" Else NomeTabela = "Compras_fornecedores"
            End If
            Set TBTransporte = CreateObject("adodb.recordset")
            TBTransporte.Open "Select * from " & NomeTabela & " where " & NomeCampo & " = " & TBAbrir!Idtransporte, Conexao, adOpenKeyset, adLockOptimistic
            If TBTransporte.EOF = False Then
                With cmbtransporte
                    Select Case TBAbrir!Tipo_transp
                        Case "C":
                            .Text = TBTransporte!NomeRazao
                            .ItemData(.NewIndex) = TBTransporte!IDCliente
                        Case "F":
                            .Text = TBTransporte!Nome_Razao
                            .ItemData(.NewIndex) = TBTransporte!IDCliente
                        Case "E":
                            .Text = TBTransporte!Empresa
                            .ItemData(.NewIndex) = TBTransporte!CODIGO
                    End Select
                End With
            End If
            TBTransporte.Close
        End If
    End If
    TBAbrir.Close
End If
If TXTIDLista <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido_lista where idlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then cmbun.Text = TBAbrir!Un
        If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com.Text = TBAbrir!Unidade_com
        If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia.Text = TBAbrir!Familia
    End If
    TBAbrir.Close
End If
If txtIDLista_serv <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido_lista where idlista = " & txtIDLista_serv, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then cmbUn_serv.Text = TBAbrir!Un
        If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com_serv.Text = TBAbrir!Unidade_com
        If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbFamilia_serv.Text = TBAbrir!Familia
    End If
    TBAbrir.Close
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro
  frmCompras_pedido_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

If txtprazo_item <> "__/__/____" Then
    VerifData = txtprazo_item
    ProcVerificaData
    If VerifData = False Then
        txtprazo_item = "__/__/____"
        txtprazo_item.SetFocus
        Exit Sub
    End If
End If
Faturamento = False
Compras_Pedido = True
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
Sit_Data = 1
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Imgcalendario2_Click()
On Error GoTo tratar_erro

If txtPrazo_serv <> "__/__/____" Then
    VerifData = txtPrazo_serv
    ProcVerificaData
    If VerifData = False Then
        txtPrazo_serv = "__/__/____"
        txtPrazo_serv.SetFocus
        Exit Sub
    End If
End If
Faturamento = False
Compras_Pedido = True
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = False
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
Sit_Data = 2
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With txtPedido
    .Locked = True
    .TabStop = False
End With
ProcLimpar
Novo_PC = True
Frame1(16).Enabled = True
Frame1(4).Enabled = True
txtResponsavel = pubUsuario
txtData = Format(Date, "dd/mm/yy")
txtStatus = "AGUARDANDO APROVAÇÃO"
cmdAdicionarfornecedor_Click
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

ProcLimparComercial
ProcLimpaCamposItem True
ProcLimpaCamposCusto
SSTab2.Tab = 0
ProcLimpaCamposServ True
ProcLimpaCamposCustoServ
ProcLimpaTotaisPedido
SSTab3.Tab = 0
txtEscopo = ""
Novo_PC1 = False
Novo_PC1_Custo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoTab_Produto()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0: Nome_anexo = "produto"
    Case 1: Nome_anexo = "centro de custo"
    Case 2: Nome_anexo = "empenho"
End Select
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", Nome_anexo, "criar novo", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", Nome_anexo, "criar novo", True, True) = False Then Exit Sub
End If
If SSTab2.Tab <> 2 Then If FunVerifSatus("criar novo " & Nome_anexo, True) = False Then Exit Sub

Select Case SSTab2.Tab
    Case 0:
        TXTIDLista = 0
        Novo_PC1 = True
        ProcLimpaCamposItem True
        Frame1(12).Enabled = True
        ProcDesbloqueiaCamposItem
        txtNomenclatura.SetFocus
        ProcLiberaTabsProd
    Case 1:
        'Verifica se o produto controla estoque e não permitie adicionar o centro de custo
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select codproduto from projproduto where desenho = '" & txtNomenclatura.Text & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            USMsgBox ("Não é permitido criar centro de custo para este produto, pois o mesmo movimenta estoque."), vbExclamation, "CAPRIND v5.0"
            TBProduto.Close
            Exit Sub
        End If
        TBProduto.Close
        
        ProcLimpaCamposCusto
        Frame1(13).Enabled = True
        Novo_PC1_Custo = True
        Cmb_centro.SetFocus
    Case 2:
        'Verifica se o produto é remessa
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select IDlista from Compras_pedido_lista where IDlista = " & TXTIDLista & " and Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            USMsgBox ("Não é permitido empenhar este produto, pois o mesmo é uma remessa."), vbExclamation, "CAPRIND v5.0"
            TBProduto.Close
            Exit Sub
        End If
        TBProduto.Close
        
        Sit_REG = 0
        Compras_Requisicao = False
        Compras_Cotacao = False
        Compras_Pedido = True
        frmProd_Lista_Produto.Show 1
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
  
If Novo_PC = True Then
    If USMsgBox("O pedido de compra ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_PC = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PC1 = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcgravarItem
        If Novo_PC1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PC1_Custo = True Then
    If USMsgBox("O centro de custo do produto não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarItem_custo
        If Novo_PC1_Custo = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PC2 = True Then
    If USMsgBox("O serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarServ
        If Novo_PC2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PC2_Custo = True Then
    If USMsgBox("O centro de custo do serviço não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarServ_custo
        If Novo_PC2_Custo = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PC3 = True Then
    If USMsgBox("O escopo de fornecimento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarEscopo
        If Novo_PC3 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_PC = False
Novo_PC1 = False
Novo_PC1_Custo = False
Novo_PC2 = False
Novo_PC2_Custo = False
Novo_PC3 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(16).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtIDfornecedor = "" Or txtFornecedor = "" Then
    NomeCampo = "o fornecedor"
    ProcVerificaAcao
    cmdAdicionarfornecedor_Click
    Exit Sub
End If
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select Pessoa, idTipoEmpresa, Data_venc, Fornecedor, Simples, Presumido, Real FROM Compras_fornecedores WHERE idcliente = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    With Cmb_empresa
        If FunVerifValidadeCertForn(.ItemData(.ListIndex), txtData, True) = False Then Exit Sub
        If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
            If FunVerifRegimeTribCliForn(.ItemData(.ListIndex), True, True) = False Then Exit Sub
        End If
    End With
End If
TBFornecedor.Close

Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select * from compras_pedido where idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = True Then
    txtPedido = FunCriarNovoNumero
    TBCompras_Pedido.AddNew
    TBCompras_Pedido!Pedido = txtPedido
    TBCompras_Pedido!Responsavel = pubUsuario
    TBCompras_Pedido!Data = Date
    TBCompras_Pedido!Status_pedido = "AGUARDANDO APROVAÇÃO"
    ProcEnviaDados
Else
    If TBCompras_Pedido!Status_pedido = "ABERTO" And txtStatus = "COMPRADO" Or TBCompras_Pedido!Status_pedido = "ABERTO" And txtStatus = "APROVADO" Then
             If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "mesmo", "este pedido", "alterar", True, True) = False Then Exit Sub
    Else
        If txtResponsavel_aprovacao <> "" Then
            If txtResponsavel_aprovacao <> pubUsuario Then If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "mesmo", "este pedido", "alterar", True, True) = False Then Exit Sub
        Else
            If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "mesmo", "este pedido", "alterar", True, True) = False Then Exit Sub
        End If
        If FunVerifSatus("alterar este pedido", True) = False Then Exit Sub
    End If
        
    If txtPedido = "" Then
        NomeCampo = "o número do pedido"
        ProcVerificaAcao
        txtPedido.SetFocus
        Exit Sub
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDPedido from compras_pedido where idpedido <> " & txtIDPedido & " and Pedido = '" & txtPedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Este número de pedido já está sendo utilizado, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtPedido.SetFocus
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
    
    MsgTexto = ""
    If TBCompras_Pedido!Status_pedido = "APROVADO" And txtStatus = "COMPRADO" Or TBCompras_Pedido!Status_pedido = "ABERTO" And txtStatus = "APROVADO" Then
        TBCompras_Pedido!Status_pedido = IIf(txtStatus = "COMPRADO", "ABERTO", txtStatus)
        Conexao.Execute "UPDATE Compras_pedido_lista Set Status_item = '" & IIf(txtStatus = "COMPRADO", "N_RECEBIDO", txtStatus) & "' where IDpedido = " & txtIDPedido & " and Status_item <> 'CANCELADO'"
        MsgTexto = " do status do pedido"
    Else
        ProcEnviaDados
    End If
End If
TBCompras_Pedido.Update
txtIDPedido = TBCompras_Pedido!IDpedido
TBCompras_Pedido.Close

If Novo_PC = True Then
    ProcGravarDCForn
    
    USMsgBox ("Novo pedido de compra cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Pedido_Localizar = "Select CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado, CP.dbl_valor_total from Compras_pedido CP LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao where CP.IDpedido = " & txtIDPedido
    ProcAtualizalistapedido (1)
Else
    USMsgBox ("Alteração " & MsgTexto & " efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(lblPaginas(3).Caption, Len(lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(3).Caption, Len(lblPaginas(3).Caption) - 5))))
    If CodigoLista <> 0 And listapedido.ListItems.Count <> 0 Then
        listapedido.SelectedItem = listapedido.ListItems(CodigoLista)
        listapedido.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Compras/Pedido"
    ID_documento = txtIDPedido
    Documento = "Nº pedido: " & txtPedido.Text
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_PC = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro

txtIDPedido = 0
txtPedido = ""
Chk_email_enviado.Value = 0
txtIDfornecedor = ""
txtResponsavel = pubUsuario
txtStatus = "AGUARDANDO APROVAÇÃO"
txtContato = ""
txtendereco = ""
txtTipo_endereco = ""
txtTipo_bairro = ""
txtNumero = ""
txtEmail = ""
txtNomenclatura = ""
txtQuantidade = ""
txtQuantidade_est = ""
txtQuantidade_PC = ""
txtEspecificacoes = ""
txtDescricao_comercial = ""
txttelefone = ""
txtCidade = ""
txtFornecedor = ""
txtuf = ""
txtCategoria = ""
txtFax = ""
Txt_n_referencia = ""
Txt_descricao_referencia = ""
txtBairro = ""
txtIPI = ""
txtvalorunitario = ""
txtData_aprovacao = ""
txtResponsavel_aprovacao = ""
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtvalorunitario = ""
TXTIDLista = "0"
txtIDLista_serv = "0"
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
cmbfamilia.ListIndex = -1
txtData = Format(Date, "dd/mm/yy")
Chk_CFOP_prod.Value = 0
CodigoLista = 0
Caption = "Administrativo - Compras - Pedido"
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaTotaisPedido()
On Error GoTo tratar_erro

txt_vlrtotalprod = "0,00"
txttotalservicos = "0,00"
txtTotaldesconto = "0,00"
txt_TotalIPI = "0,00"
txt_ICMS_ST = "0,00"
txt_BaseICMS = "0,00"
txt_vlrICMS = "0,00"
txt_baseICMS_ST = "0,00"
TxtTotalFrete = "0,00"
txtTotalSeguro = "0,00"
TxtTotalacessorias = "0,00"
txtTotalPedido = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposBanco()
On Error GoTo tratar_erro

TBCompras_Lista!IDpedido = 0
TBCompras_Lista!Quant_Comp = 0
TBCompras_Lista!preco_unitario = 0
TBCompras_Lista!IPI = 0
TBCompras_Lista!preco_total = 0
TBCompras_Lista!Nota_fiscal = 0
TBCompras_Lista!Data_emissao = Null
TBCompras_Lista!vlrICMS = 0
TBCompras_Lista!VlrIPI = 0
TBCompras_Lista!ICMS = 0
TBCompras_Lista!Total_NF = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Label1_DblClick(index As Integer)
On Error GoTo tratar_erro

    If Novo_PC = True Then Exit Sub
    
    If TXTIDLista <> 0 Or txtIDLista_serv <> 0 Then
        If TXTIDLista <> 0 Then
            IDlista = TXTIDLista
        Else
            IDlista = txtIDLista_serv
        End If
        frmCompras_pedido_liberar.Show 1
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_custo
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtResponsavel_aprovacao <> "" Then
                    If txtResponsavel_aprovacao <> pubUsuario Then If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                Else
                    If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                End If
                If FunVerifSatus("", False) = False Then GoTo Proximo
                If FunVerifSatusProdServ(txtstatus_item, "", False, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_custo, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_custo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If txtResponsavel_aprovacao <> "" Then
                If txtResponsavel_aprovacao <> pubUsuario Then
                    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "excluir", True, True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
            Else
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "excluir", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            If FunVerifSatus("excluir este centro de custo", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifSatusProdServ(txtstatus_item, "excluir este centro de custo", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custo_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_custo.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CPLC.*, US.Codigo, US.Setor, US.DtBloq, US.ID AS ID_centro from Compras_pedido_lista_custo CPLC INNER JOIN Usuarios_setor US ON CPLC.ID_CC = US.ID where CPLC.id = " & Lista_custo.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposCusto
    txtIDCentro = TBAbrir!ID
    
    With Cmb_centro
        If IsNull(TBAbrir!CODIGO) = False And TBAbrir!CODIGO <> "" Then
            If IsNull(TBAbrir!DtBloq) = False Then
                .AddItem TBAbrir!CODIGO & " - " & IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
                .ItemData(Cmb_centro.NewIndex) = TBAbrir!ID_centro
            End If
            .Text = TBAbrir!CODIGO & " - " & IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
        Else
            If IsNull(TBAbrir!DtBloq) = False Then
                .AddItem IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
                .ItemData(Cmb_centro.NewIndex) = TBAbrir!ID_centro
            End If
            .Text = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
        End If
    End With
    
    txtValorCentro = IIf(IsNull(TBAbrir!valor), "", Format(TBAbrir!valor, "###,##0.00"))
    txtPercentualCentro = IIf(IsNull(TBAbrir!Percentual), "", Format(TBAbrir!Percentual, "###,##0.0000000000"))
    CodigoLista2 = Lista_custo.SelectedItem.index
End If
TBAbrir.Close
Frame1(13).Enabled = True
Novo_PC1_Custo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custoServ_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_custoServ
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtResponsavel_aprovacao <> "" Then
                    If txtResponsavel_aprovacao <> pubUsuario Then If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                Else
                    If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                End If
                If FunVerifSatus("", False) = False Then GoTo Proximo
                If FunVerifSatusProdServ(txtStatus_serv, "", False, False) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_custoServ, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custoServ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_custoServ
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If txtResponsavel_aprovacao <> "" Then
                If txtResponsavel_aprovacao <> pubUsuario Then
                    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "excluir", True, True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
            Else
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "excluir", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            If FunVerifSatus("excluir este centro de custo", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifSatusProdServ(txtStatus_serv, "excluir este centro de custo", True, False) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custoServ_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_custoServ.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CPLC.*, US.Codigo, US.Setor, US.DtBloq, US.ID AS ID_centro from Compras_pedido_lista_custo CPLC INNER JOIN Usuarios_setor US ON CPLC.ID_CC = US.ID where CPLC.id = " & Lista_custoServ.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposCustoServ
    txtIDCentro_serv = TBAbrir!ID
    
    With Cmb_centro_servico
        If IsNull(TBAbrir!CODIGO) = False And TBAbrir!CODIGO <> "" Then
            If IsNull(TBAbrir!DtBloq) = False Then
                .AddItem TBAbrir!CODIGO & " - " & IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
                .ItemData(Cmb_centro_servico.NewIndex) = TBAbrir!ID_centro
            End If
            .Text = TBAbrir!CODIGO & " - " & IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
        Else
            If IsNull(TBAbrir!DtBloq) = False Then
                .AddItem IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
                .ItemData(Cmb_centro_servico.NewIndex) = TBAbrir!ID_centro
            End If
            .Text = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
        End If
    End With
    
    txtValorCentro_Serv = IIf(IsNull(TBAbrir!valor), "", Format(TBAbrir!valor, "###,##0.00"))
    txtPercentualCentro_Serv = IIf(IsNull(TBAbrir!Percentual), "", Format(TBAbrir!Percentual, "###,##0.0000000000"))
    CodigoLista4 = Lista_custoServ.SelectedItem.index
End If
TBAbrir.Close
Frame1(8).Enabled = True
Novo_PC2_Custo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_empenhos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifSatus("", False) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_empenhos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_DblClick()
On Error GoTo tratar_erro

Qtde = 0
qtde_solicitada = ""
With Lista_empenhos
    If .ListItems.Count = 0 Then Exit Sub
    If .SelectedItem.ListSubItems(16) = "FATURADO" Or .SelectedItem.ListSubItems(16) = "FATURADO PARCIAL" Then
        If USMsgBox("Deseja alterar a quantidade empenhada?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If Alterar = False Then
                USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
Mensagem:
            qtde_solicitada = txtQuantidade
            qtde_solicitada = InputBox("Favor informar a quantidade empenhada.", , qtde_solicitada)
            If qtde_solicitada = "" Then Exit Sub
            
            If IsNumeric(qtde_solicitada) = False Then
                USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            Qtde = qtde_solicitada
            If Qtde <= 0 Then
                USMsgBox ("So é permitido quantidade maior que 0."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            
            'Verifica se a quantidade empenhada é maior que a quantidade solicitada
            Qtd = txtQS
            valor = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Qtde_empenho) as Valor from Compras_pedido_lista_empenhos where IDcarteira = " & .SelectedItem.ListSubItems(1) & " and ID <> " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            If Qtd < (valor + Qtde) Then
                USMsgBox ("A quantidade empenhada não pode ser maior que a quantidade comprada, favor alterar."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
                        
            NovoValor = Replace(Qtde, ",", ".")
            Conexao.Execute "Update Compras_pedido_lista_empenhos Set Qtde_empenho = " & NovoValor & " where ID = " & .SelectedItem
            
            USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Empenhar produto"
            ID_documento = .SelectedItem
            Documento = "Nº pedido: " & txtPedido & " - Cód. interno: " & txtNomenclatura
            Documento1 = "Pedido int.: " & .SelectedItem.ListSubItems(2) & " - Rev.: " & .SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(4) & " - Rev.: " & .SelectedItem.ListSubItems(3) & " - Qtde. empenhada: " & .SelectedItem.ListSubItems(9) & " - Qtde. entrada: " & .SelectedItem.ListSubItems(10)
            ProcGravaEvento
            '==================================
            ProcCarregaListaEmpenhosProd
        Else
            ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), .SelectedItem.ListSubItems(5), False
        End If
    Else
        ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), .SelectedItem.ListSubItems(5), False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_empenhos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifSatus("excluir este empenho", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Lista_empenhos_serv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_empenhos_serv
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerifSatus("", False) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_empenhos_serv, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_serv_DblClick()
On Error GoTo tratar_erro

Qtde = 0
qtde_solicitada = ""
With Lista_empenhos_serv
    If .ListItems.Count = 0 Then Exit Sub
    If .SelectedItem.ListSubItems(16) = "FATURADO" Or .SelectedItem.ListSubItems(16) = "FATURADO PARCIAL" Then
        If USMsgBox("Deseja alterar a quantidade empenhada?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If Alterar = False Then
                USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
Mensagem:
            qtde_solicitada = txtQtde_serv
            qtde_solicitada = InputBox("Favor informar a quantidade empenhada.", , qtde_solicitada)
            If qtde_solicitada = "" Then Exit Sub
            
            If IsNumeric(qtde_solicitada) = False Then
                USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            Qtde = qtde_solicitada
            If Qtde <= 0 Then
                USMsgBox ("So é permitido quantidade maior que 0."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            
            'Verifica se a quantidade empenhada é maior que a quantidade solicitada
            Qtd = txtQS
            valor = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Qtde_empenho) as Valor from Compras_pedido_lista_empenhos where IDcarteira = " & .SelectedItem.ListSubItems(1) & " and ID <> " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            If Qtd < (valor + Qtde) Then
                USMsgBox ("A quantidade empenhada não pode ser maior que a quantidade comprada, favor alterar."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
                        
            NovoValor = Replace(Qtde, ",", ".")
            Conexao.Execute "Update Compras_pedido_lista_empenhos Set Qtde_empenho = " & NovoValor & " where ID = " & .SelectedItem
            
            USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Empenhar serviço"
            ID_documento = .SelectedItem
            Documento = "Nº pedido: " & txtPedido & " - Cód. interno: " & txtCodigo
            Documento1 = "Pedido int.: " & .SelectedItem.ListSubItems(2) & " - Rev.: " & .SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(4) & " - Rev.: " & .SelectedItem.ListSubItems(3) & " - Qtde. empenhada: " & .SelectedItem.ListSubItems(9) & " - Qtde. entrada: " & .SelectedItem.ListSubItems(10)
            ProcGravaEvento
            '==================================
            ProcCarregaListaEmpenhosServ
        Else
            ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), .SelectedItem.ListSubItems(5), False
        End If
    Else
        ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), .SelectedItem.ListSubItems(5), False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_serv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_empenhos_serv
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerifSatus("excluir este empenho", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_solicitados
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_solicitados, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNecessidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaNecessidade
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaNecessidade, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosFabricante()
On Error GoTo tratar_erro

ListaFabricante.ListItems.Clear
StrSql = "SELECT TOP (100) PERCENT PFAB.Part_number, FM.Fabricante FROM Fabricante_marca AS FM RIGHT OUTER JOIN Projproduto_fabricante AS PFAB ON FM.Id = PFAB.Idfabricante RIGHT OUTER JOIN Projproduto AS PP ON PFAB.Codproduto = PP.codproduto WHERE PP.Desenho = '" & ListaNecessidade.SelectedItem.ListSubItems.Item(1).Text & "'"
Contador = 1
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
 If TBEstoque.EOF = False Then
     Do While TBEstoque.EOF = False
        With ListaFabricante.ListItems
            .Add , , Contador
            .Item(.Count).SubItems(1) = IIf(IsNull(TBEstoque!Part_number), "", TBEstoque!Part_number)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBEstoque!Fabricante), "", TBEstoque!Fabricante)
        End With
        Contador = Contador + 1
        TBEstoque.MoveNext
    Loop
 End If
 TBEstoque.Close
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaNecessidade_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaNecessidade.ListItems.Count = 0 Then Exit Sub

ProcCarregaDadosFabricante

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub listapedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With listapedido
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ID_empresa, Status_pedido, DtValidacao from Compras_pedido where IDpedido = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    status = FunCorrigeStatusPedido(TBFI!Status_pedido)
                    If Cmb_opcao_lista = "Validação" Then
                        If status <> "AGUARDANDO APROVAÇÃO" And status <> "CANCELADO" Then GoTo Proximo
                        
                        'Verifica se algum item não foi informado o centro de custo
                        If IsNull(TBFI!DtValidacao) = True Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select codigo from Empresa where Codigo = " & TBFI!ID_empresa & " and CC_obrigatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                Set TBCompras_Pedido = CreateObject("adodb.recordset")
                                TBCompras_Pedido.Open "Select CPL.idlista from compras_pedido_lista CPL LEFT JOIN Projproduto P ON P.Desenho = CPL.Desenho where CPL.Idpedido = " & .ListItems(InitFor) & " and (P.Estoque = 'False' or P.Estoque IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
                                If TBCompras_Pedido.EOF = False Then
                                    Do While TBCompras_Pedido.EOF = False
                                        Set TBContas = CreateObject("adodb.recordset")
                                        TBContas.Open "Select idlista from compras_pedido_lista_custo where idlista = " & TBCompras_Pedido!IDlista, Conexao, adOpenKeyset, adLockOptimistic
                                        If TBContas.EOF = True Then
                                            TBContas.Close
                                            TBCompras_Pedido.Close
                                            TBAbrir.Close
                                            GoTo Proximo
                                        End If
                                        TBContas.Close
                                        TBCompras_Pedido.MoveNext
                                    Loop
                                End If
                                TBCompras_Pedido.Close
                            End If
                            TBAbrir.Close
                            
                            'Verifica se tem algum produto/serviço sem prazo final
                            Set TBCompras_Pedido = CreateObject("adodb.recordset")
                            TBCompras_Pedido.Open "Select idlista from compras_pedido_lista where Idpedido = " & .ListItems(InitFor) & " and Prazo IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCompras_Pedido.EOF = False Then
                                TBCompras_Pedido.Close
                                GoTo Proximo
                            End If
                            TBCompras_Pedido.Close
                        End If
                    Else
                        If status <> "AGUARDANDO APROVAÇÃO" And status <> "APROVADO" And status <> "COMPRADO" Then GoTo Proximo
                        If IsNull(TBFI!DtValidacao) = True Then GoTo Proximo
                        
                        'Verifica se o usuario pode aprovar o pedido de acordo com o limite cadastrado
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select A.Valor_Limite from usuarios U INNER JOIN acessos A ON A.IDUsuario = U.IDUsuario where U.usuario = '" & pubUsuario & "' and A.Acesso = 'Compras/Pedido/Aprovar' and A.Validacao = 'True' and Valor_Limite IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Set TBCompras_Pedido = CreateObject("adodb.recordset")
                            TBCompras_Pedido.Open "Select dbl_valor_total from compras_pedido where Idpedido = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                            If TBCompras_Pedido.EOF = False Then
                                Qtde = IIf(IsNull(TBCompras_Pedido!dbl_valor_total), 0, TBCompras_Pedido!dbl_valor_total)
                                Qtd = TBAbrir!Valor_Limite
                                If Qtde > Qtd Then
                                    TBAbrir.Close
                                    TBCompras_Pedido.Close
                                    GoTo Proximo
                                End If
                            End If
                            TBCompras_Pedido.Close
                        End If
                        TBAbrir.Close
                    End If
                End If
                TBFI.Close
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView listapedido, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listapedido_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With listapedido
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select ID_empresa, Status_pedido, DtValidacao from Compras_pedido where IDpedido = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                status = FunCorrigeStatusPedido(TBFI!Status_pedido)
                If Cmb_opcao_lista = "Validação" Then
'                    If status <> "AGUARDANDO APROVAÇÃO" And status <> "CANCELADO" Then
'                        USMsgBox ("Não é permitido validar/cancelar validação, pois o pedido de compra esta " & status & "."), vbExclamation, "CAPRIND v5.0"
'                        .ListItems.Item(InitFor).Checked = False
'                        Exit Sub
'                    End If
                    
                    If IsNull(TBFI!DtValidacao) = True Then
                        'Verifica se algum item não foi informado o centro de custo
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select codigo from Empresa where Codigo = " & TBFI!ID_empresa & " and CC_obrigatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Set TBCompras_Pedido = CreateObject("adodb.recordset")
                            TBCompras_Pedido.Open "Select CPL.idlista from compras_pedido_lista CPL LEFT JOIN Projproduto P ON P.Desenho = CPL.Desenho where CPL.Idpedido = " & .ListItems(InitFor) & " and (P.Estoque = 'False' or P.Estoque IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCompras_Pedido.EOF = False Then
                                Do While TBCompras_Pedido.EOF = False
                                    Set TBContas = CreateObject("adodb.recordset")
                                    TBContas.Open "Select idlista from compras_pedido_lista_custo where idlista = " & TBCompras_Pedido!IDlista, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBContas.EOF = True Then
                                        If USMsgBox("Existe(m) produto(s)/serviço(s) sem centro de custo cadastrado, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                                            .ListItems.Item(InitFor).Checked = False
                                            TBContas.Close
                                            TBCompras_Pedido.Close
                                            TBAbrir.Close
                                            Exit Sub
                                        Else
                                            TBContas.Close
                                            GoTo Proximo
                                        End If
                                    End If
                                    TBCompras_Pedido.MoveNext
                                Loop
                            End If
Proximo:
                            TBCompras_Pedido.Close
                        End If
                        TBAbrir.Close
                        
                        'Verifica se tem algum produto/serviço sem prazo final
                        Set TBCompras_Pedido = CreateObject("adodb.recordset")
                        TBCompras_Pedido.Open "Select idlista from compras_pedido_lista where Idpedido = " & .ListItems(InitFor) & " and Prazo IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras_Pedido.EOF = False Then
                            USMsgBox ("Não é permitido validar este pedido, pois existe(m) produto(s)/serviço(s) sem prazo final cadastrado."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                        End If
                        TBCompras_Pedido.Close
                    End If
                Else
                    If status <> "AGUARDANDO APROVAÇÃO" And status <> "APROVADO" And status <> "COMPRADO" And status <> "RECEBIDO PARCIAL" Then
                        USMsgBox ("Não é permitido aprovar/cancelar aprovação, pois o pedido de compra esta " & status & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If IsNull(TBFI!DtValidacao) = True Then
                        USMsgBox ("Não é permitido aprovar este pedido de compra, pois o mesmo ainda não foi validado."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    
                    'Verifica se o usuario pode aprovar o pedido de acordo com o limite cadastrado
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select A.Valor_Limite from usuarios U INNER JOIN acessos A ON A.IDUsuario = U.IDUsuario where U.usuario = '" & pubUsuario & "' and A.Acesso = 'Compras/Pedido/Aprovar' and A.Validacao = 'True' and Valor_Limite IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Set TBCompras_Pedido = CreateObject("adodb.recordset")
                        TBCompras_Pedido.Open "Select dbl_valor_total from compras_pedido where Idpedido = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBCompras_Pedido.EOF = False Then
                            Qtde = IIf(IsNull(TBCompras_Pedido!dbl_valor_total), 0, TBCompras_Pedido!dbl_valor_total)
                            Qtd = TBAbrir!Valor_Limite
                            If Qtde > Qtd Then
                                USMsgBox ("Atenção usuário " & pubUsuario & ", você não tem autorização para aprovar este pedido pois ultrapassou o valor permitido."), vbExclamation, "CAPRIND v5.0"
                                .ListItems.Item(InitFor).Checked = False
                                TBAbrir.Close
                                TBCompras_Pedido.Close
                                Exit Sub
                            End If
                        End If
                        TBCompras_Pedido.Close
                    End If
                    TBAbrir.Close
                End If
            End If
            TBFI.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listapedido_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listapedido.ListItems.Count = 0 Or IsNumeric(listapedido.SelectedItem) = False Then Exit Sub

Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select * from compras_pedido where idpedido = " & listapedido.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    ProcLimpar
    ProcLimpaCamposItem True
    ProcLimpaCamposServ True
    ProcPuxaDados
    CodigoLista = listapedido.SelectedItem.index
    txtIDPedido = listapedido.SelectedItem
    
  If txtStatus.Text = "AGUARDANDO APROVAÇÃO" And txtDtValidacao.Text <> "" And txtIDPedido <> 0 Then
'    ProcBuscarPedidoWEB (txtIDPedido)
    ProcSalvarPedidoWEB (Int(txtIDPedido))
  End If

End If
TBCompras_Pedido.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaServ_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaServ
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtResponsavel_aprovacao <> "" Then
                    If txtResponsavel_aprovacao <> pubUsuario Then
                        If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                    End If
                Else
                    If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                End If
                If FunVerifSatus("", False) = False Then GoTo Proximo
                If FunVerifSatusProdServ(.ListItems(InitFor).SubItems(11), "", False, False) = False Then GoTo Proximo
                
                Set TBCompras_Lista = CreateObject("adodb.recordset")
                TBCompras_Lista.Open "Select IDlista from compras_pedido_lista where idlista = " & .ListItems(InitFor) & " and ID_programacao <> 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBCompras_Lista.EOF = False Then
                    TBCompras_Lista.Close
                    GoTo Proximo
                End If
                TBCompras_Lista.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaServ, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaServ_DblClick()
On Error GoTo tratar_erro

With ListaServ
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturada .SelectedItem, .SelectedItem.ListSubItems(1)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaServ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaServ
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If txtResponsavel_aprovacao <> "" Then
                If txtResponsavel_aprovacao <> pubUsuario Then
                    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este serviço", "excluir", True, True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
            Else
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este serviço", "excluir", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            If FunVerifSatus("excluir este serviço", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifSatusProdServ(.ListItems(InitFor).SubItems(11), "excluir este serviço", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select idlista from compras_pedido_lista where idlista = " & .ListItems(InitFor) & " and ID_programacao <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then
                USMsgBox ("Não é permitido excluir este serviço, pois o mesmo está vinculado a uma programação."), vbExclamation, "CAPRIND v5.0"
                TBCompras_Lista.Close
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            TBCompras_Lista.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaServ_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaServ.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from compras_pedido_lista where idlista = " & ListaServ.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtIDLista_serv = 0
    Frame1(7).Enabled = True
    ProcLimpaCamposServ True
    ProcPuxaDadosListaServ
    CodigoLista3 = ListaServ.SelectedItem.index
End If
TBProduto.Close
Novo_PC2 = False
ProcLiberaTabsServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Listprod
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtResponsavel_aprovacao <> "" Then
                    If txtResponsavel_aprovacao <> pubUsuario Then
                        If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                    End If
                Else
                    If FunVerificaRegistroValidadoSemMsg("Compras_pedido", "IDPedido = " & txtIDPedido, True) = False Then GoTo Proximo
                End If
                If FunVerifSatus("", False) = False Then GoTo Proximo
                If FunVerifSatusProdServ(.ListItems(InitFor).SubItems(13), "", False, True) = False Then GoTo Proximo
                
                Set TBCompras_Lista = CreateObject("adodb.recordset")
                TBCompras_Lista.Open "Select IDlista from compras_pedido_lista where idlista = " & .ListItems(InitFor) & " and ID_programacao <> 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBCompras_Lista.EOF = False Then
                    TBCompras_Lista.Close
                    GoTo Proximo
                End If
                TBCompras_Lista.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Listprod, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_DblClick()
On Error GoTo tratar_erro

With Listprod
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturada .SelectedItem, .SelectedItem.ListSubItems(1)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifQtdeFaturada(IDlista As Long, Codinterno As String)
On Error GoTo tratar_erro

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_pedido_lista where IDlista = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    TextoNF = ""
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from Estoque_controle_recebimento where IDpedido = " & txtIDPedido & " and IdLista = " & IDlista & " and Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Do While TBEstoque.EOF = False
            If TextoNF = "" Then
                TextoNF = "NF: " & TBEstoque!Nota_fiscal & " - Dt. emissão: " & Format(TBEstoque!Data_emissao, "dd/mm/yy") & " - Vlr. total: " & Format(TBEstoque!Total_NF, "###,##0.00") & " - Qtde.: " & Format(TBEstoque!Recebido, "###,##0.0000") & " " & TBCompras!Unidade_com & " - Certificado: " & TBEstoque!Certificado & " - Corrida: " & TBEstoque!Corrida
            Else
                TextoNF = TextoNF & "  |  NF: " & TBEstoque!Nota_fiscal & " - Dt. emissão: " & Format(TBEstoque!Data_emissao, "dd/mm/yy") & " - Vlr. total: " & Format(TBEstoque!Total_NF, "###,##0.00") & " - Qtde.: " & Format(TBEstoque!Recebido, "###,##0.0000") & " " & TBCompras!Unidade_com & " - Certificado: " & TBEstoque!Certificado & " - Corrida: " & TBEstoque!Corrida
            End If
            TBEstoque.MoveNext
        Loop
    Else
        CamposFiltro = "NF.TipoNF, NF.int_NotaFiscal, NF.dt_dataemissao, NFPP.Quantidade"
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select " & CamposFiltro & " from tbl_dados_nota_fiscal NF INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_nota = NF.ID where NF.int_TipoNota = 2 and NF.Pedido_interno = 'False' and NFPP.ID_carteira = " & IDlista & " and NFPP.Codinterno = '" & Codinterno & "' group by " & CamposFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Do While TBEstoque.EOF = False
                Select Case TBEstoque!TipoNF
                    Case "M1": Tipo = "Produto(s)"
                    Case "SA": Tipo = "Serviço(s)"
                    Case "M1SA": Tipo = "Produto(s)/Serviço(s)"
                End Select
                If TextoNF = "" Then
                    TextoNF = "NF: " & TBEstoque!int_NotaFiscal & " - Tipo: " & Tipo & " - Dt. emissão : " & Format(TBEstoque!dt_DataEmissao, "dd/mm/yy") & " - Qtde. : " & Format(TBEstoque!quantidade, "###,##0.0000")
                Else
                    TextoNF = TextoNF & "  |  NF: " & TBEstoque!int_NotaFiscal & " - Tipo: " & Tipo & " - Dt. emissão : " & Format(TBEstoque!dt_DataEmissao, "dd/mm/yy") & " - Qtde. : " & Format(TBEstoque!quantidade, "###,##0.0000")
                End If
                TBEstoque.MoveNext
            Loop
        End If
    End If
    TBEstoque.Close
    
    If TBCompras!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBCompras!Status_Item = "RECEBIDO" Or TBCompras!Status_Item = "CANCELADO" Then
        Status_Item = TBCompras!Status_Item
    ElseIf TBCompras!Status_Item = "N_RECEBIDO" Then
            Status_Item = "COMPRADO"
        Else
            Status_Item = "RECEBIDO PARCIAL"
    End If
    USMsgBox ("Cód. interno: " & TBCompras!Desenho & " " & vbCrLf & "Status: " & Status_Item & " " & vbCrLf & " " & TextoNF), vbInformation, "CAPRIND v5.0"
End If
TBCompras.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Listprod
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If txtResponsavel_aprovacao <> "" Then
                If txtResponsavel_aprovacao <> pubUsuario Then
                    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este produto", "excluir", True, True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
            Else
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este produto", "excluir", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            If FunVerifSatus("excluir este produto", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerifSatusProdServ(.ListItems(InitFor).SubItems(13), "excluir este produto", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select idlista from compras_pedido_lista where idlista = " & .ListItems(InitFor) & " and ID_programacao <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then
                USMsgBox ("Não é permitido excluir este produto, pois o mesmo está vinculado a uma programação."), vbExclamation, "CAPRIND v5.0"
                TBCompras_Lista.Close
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            TBCompras_Lista.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Listprod.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from compras_pedido_lista where idlista = " & Listprod.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TXTIDLista = 0
    ProcLimpaCamposItem True
    Frame1(12).Enabled = True
    ProcPuxaDadosLista
    CodigoLista1 = Listprod.SelectedItem.index
End If
TBProduto.Close
Novo_PC1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposItem()
On Error GoTo tratar_erro

txtNomenclatura.Locked = True
cmdfiltrar.Enabled = False
CmdEscolher_item.Enabled = False
Frame1(14).Enabled = False
cmbReferencia.Locked = True
txtEspecificacoes.Locked = True
cmdCFOP_prod.Enabled = False
cmdLimpar_CFOP_prod.Enabled = False
If txtstatus_item = "RECEBIDO" Then
    framePrazo.Enabled = False
    imgCalendario.Enabled = False
End If
chkRemessa.Enabled = False
txtDescricao_comercial.Locked = True
txtObs.Locked = True
cmbfamilia.Locked = True
cmbun.Locked = True
Cmb_un_com.Locked = True
cmdCF.Enabled = False
cmdLimpar_NCM.Enabled = False
Cmb_CST_ICMS.Locked = True
txtdetalheitem.Locked = True
txtvalorunitario.Locked = True
Chk_desc.Enabled = False
Chk_valor_desc.Enabled = False
txtFrete.Locked = True
ChkFrete_IPI.Enabled = False
txtSeguro.Locked = True
txtAcessorias.Locked = True
txtQuantidade.Locked = True
txtIPI.Locked = True
txtICMS.Locked = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposServ()
On Error GoTo tratar_erro

txtCodigo.Locked = True
cmdFiltrar_codigo.Enabled = False
cmdAbrir_codigo.Enabled = False
Frame1(15).Enabled = False
cmbreferencia_serv.Locked = True
txtDescricao_serv.Locked = True
cmdCFOP_serv.Enabled = False
cmdLimpar_CFOP_serv.Enabled = False
If txtStatus_serv = "RECEBIDO" Then
    framePrazo_serv.Enabled = False
    imgCalendario2.Enabled = False
End If
txtDescricao_comercialServ.Locked = True
txtObs_serv.Locked = True
cmbFamilia_serv.Locked = True
txtDetalhe_serv.Locked = True
cmbUn_serv.Locked = True
'Cmb_un_com_serv.Locked = True
txtValorUnit_serv.Locked = True
Chk_desc2.Enabled = False
Chk_valor_desc2.Enabled = False
txtQtde_serv.Locked = True
txtISSQN.Locked = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcDesbloqueiaCamposItem()
On Error GoTo tratar_erro

txtNomenclatura.Locked = False
cmdfiltrar.Enabled = True
CmdEscolher_item.Enabled = True
Frame1(14).Enabled = True
cmbReferencia.Locked = False
txtEspecificacoes.Locked = False
cmdCFOP_prod.Enabled = True
cmdLimpar_CFOP_prod.Enabled = True
framePrazo.Enabled = True
imgCalendario.Enabled = True
chkRemessa.Enabled = True
txtDescricao_comercial.Locked = False
txtObs.Locked = False
cmbfamilia.Locked = False
cmbun.Locked = False
Cmb_un_com.Locked = False
cmdCF.Enabled = True
cmdLimpar_NCM.Enabled = True
Cmb_CST_ICMS.Locked = False
txtdetalheitem.Locked = False
txtvalorunitario.Locked = False
Chk_desc.Enabled = True
Chk_valor_desc.Enabled = True
txtFrete.Locked = False
ChkFrete_IPI.Enabled = True
txtSeguro.Locked = False
txtAcessorias.Locked = False
txtQuantidade.Locked = False
txtIPI.Locked = False
txtICMS.Locked = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposItemSolCot()
On Error GoTo tratar_erro

With txtNomenclatura
    .Locked = True
    .TabStop = False
End With
With cmbReferencia
    .Locked = True
    .TabStop = False
End With
With txtEspecificacoes
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercial
    .Locked = True
    .TabStop = False
End With
With txtdetalheitem
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposServSolCot()
On Error GoTo tratar_erro

With txtCodigo
    .Locked = True
    .TabStop = False
End With
With cmbreferencia_serv
    .Locked = True
    .TabStop = False
End With
With txtDescricao_serv
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercialServ
    .Locked = True
    .TabStop = False
End With
With txtDetalhe_serv
    .Locked = True
    .TabStop = False
End With
With cmbFamilia_serv
    .Locked = True
    .TabStop = False
End With
With cmbUn_serv
    .Locked = True
    .TabStop = False
End With
'With Cmb_un_com_serv
'    .Locked = False
'    .TabStop = False
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcDesbloqueiaCamposServ()
On Error GoTo tratar_erro

txtCodigo.Locked = False
cmdFiltrar_codigo.Enabled = True
cmdAbrir_codigo.Enabled = True
Frame1(15).Enabled = True
cmbreferencia_serv.Locked = False
txtDescricao_serv.Locked = False
cmdCFOP_serv.Enabled = True
cmdLimpar_CFOP_serv.Enabled = True
framePrazo_serv.Enabled = True
imgCalendario2.Enabled = True
txtDescricao_comercialServ.Locked = False
txtObs_serv.Locked = False
cmbFamilia_serv.Locked = False
txtDetalhe_serv.Locked = False
cmbUn_serv.Locked = False
'Cmb_un_com_serv.Locked = False
txtValorUnit_serv.Locked = False
Chk_desc2.Enabled = True
Chk_valor_desc2.Enabled = True
txtQtde_serv.Locked = False
txtISSQN.Locked = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesbloqueiaCamposItemSolCot()
On Error GoTo tratar_erro

With txtNomenclatura
    .Locked = False
    .TabStop = True
End With
With cmbReferencia
    .Locked = False
    .TabStop = True
End With
With txtEspecificacoes
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercial
    .Locked = False
    .TabStop = True
End With
With txtdetalheitem
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
With cmbun
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesbloqueiaCamposServSolCot()
On Error GoTo tratar_erro

With txtCodigo
    .Locked = False
    .TabStop = True
End With
With cmbreferencia_serv
    .Locked = False
    .TabStop = True
End With
With txtDescricao_serv
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercialServ
    .Locked = False
    .TabStop = True
End With
With txtDetalhe_serv
    .Locked = False
    .TabStop = True
End With
With cmbFamilia_serv
    .Locked = False
    .TabStop = True
End With
With cmbUn_serv
    .Locked = False
    .TabStop = True
End With
'With Cmb_un_com_serv
'    .Locked = False
'    .TabStop = True
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_PCP_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_vendas_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptFim_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

ProcCorrigeForm
Select Case SSTab1.Tab
    Case 0:
        If ListaNecessidade.Visible = True Then ListaNecessidade.SetFocus
    Case 1:
        If listapedido.Visible = True Then listapedido.SetFocus
    Case 2:
    
    If IsNumeric(txtIDPedido.Text) = False Or txtIDPedido.Text = "" Then
        USMsgBox "Escolha um pedido na lista pra visualizar os dados comerciais", vbInformation, "CAPRIND v.5"
        SSTab1.Tab = 1
        Exit Sub
    End If
    
        If FunVerificaProsseguir(False, False) = False Then Exit Sub
        cmbpagamento.SetFocus
        ProcLimparComercial
        ProcAbreComercial
    Case 3:
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposProdSemCadastrado Else ProcBloqueiaCamposProdComCadastrado
        If FunVerificaProsseguir(False, False) = False Then Exit Sub
        Listprod.SetFocus
        ProcAtualizalista
    Case 4:
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then ProcLiberaCamposServSemCadastrado Else ProcBloqueiaCamposServComCadastrado
        If FunVerificaProsseguir(False, False) = False Then Exit Sub
        ListaServ.SetFocus
        ProcAtualizalistaServ
    Case 5:
        If FunVerificaProsseguir(False, False) = False Then Exit Sub
        txtEscopo.SetFocus
        ProcCarregaEscopoForn
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaProsseguir(Produto As Boolean, Servico As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerificaProsseguir = True
If Produto = True Then
    If Novo_PC1 = True Then
        USMsgBox ("Salve o produto antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
        SSTab2.Tab = 0
        FunVerificaProsseguir = False
        Exit Function
    End If
ElseIf Servico = True Then
        If Novo_PC2 = True Then
            USMsgBox ("Salve o serviço antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab3.Tab = 0
            FunVerificaProsseguir = False
            Exit Function
        End If
    Else
        If txtIDPedido = 0 Then
            FunVerificaProsseguir = False
            SSTab1.Tab = 1
            Exit Function
        End If
        If Novo_PC = True Then
            SSTab1.Tab = 1
            USMsgBox ("Salve o pedido de compra antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            FunVerificaProsseguir = False
        End If
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If TXTIDLista = 0 Then
    SSTab2.Tab = 0
    Exit Sub
End If
With USToolBar4
    Select Case SSTab2.Tab
        Case 0:
            .ButtonState(2) = 0
            .ButtonState(5) = 0
            .ButtonState(6) = 0
            .ButtonState(7) = 0
            .ButtonState(8) = 0
            .ButtonState(9) = 0
            .ButtonState(10) = 5
        Case 1:
            .ButtonState(2) = 0
            .ButtonState(5) = 5
            .ButtonState(6) = 5
            .ButtonState(7) = 5
            .ButtonState(8) = 5
            .ButtonState(9) = 5
            .ButtonState(10) = 0
        Case 2:
            .ButtonState(2) = 5
            .ButtonState(5) = 0
            .ButtonState(5) = 5
            .ButtonState(6) = 5
            .ButtonState(7) = 5
            .ButtonState(8) = 5
            .ButtonState(9) = 5
            .ButtonState(10) = 5
        
    End Select
    .Refresh
End With

Select Case SSTab2.Tab
    Case 0: Listprod.SetFocus
    Case 1:
        Lista_custo.SetFocus
        txtVlrTotal_centro = txtvlrTotal
        If FunVerificaProsseguir(True, False) = False Then Exit Sub
        ProcCarregaLista_Custo
    Case 2:
'        Lista_empenhos.SetFocus
        If FunVerificaProsseguir(True, False) = False Then Exit Sub
        Txt_qtde_total_comprada(0) = txtQuantidade

        'Verifica se é requisição de serviço de terceiro
        If txtOrdem <> "" And txtOrdem <> "0" Then
            USMsgBox ("Não é permitido fazer o empenho, pois este produto já está empenhado para uma ordem de produção."), vbExclamation, "CAPRIND v5.0"
            SSTab2.Tab = 0
            Exit Sub
        End If
        ProcCarregaListaEmpenhosProd
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtIDLista_serv = 0 Then
    SSTab3.Tab = 0
    Exit Sub
End If
With USToolBar5
    Select Case SSTab3.Tab
        Case 0:
            .ButtonState(2) = 0
            .ButtonState(5) = 0
            .ButtonState(6) = 0
            .ButtonState(7) = 0
            .ButtonState(8) = 0
            .ButtonState(9) = 5
        Case 1:
            .ButtonState(2) = 0
            .ButtonState(5) = 5
            .ButtonState(6) = 5
            .ButtonState(7) = 5
            .ButtonState(8) = 5
            .ButtonState(9) = 0
        Case 2:
            .ButtonState(2) = 5
            .ButtonState(5) = 5
            .ButtonState(6) = 5
            .ButtonState(7) = 5
            .ButtonState(8) = 5
            .ButtonState(9) = 5
    End Select
    .Refresh
End With

Select Case SSTab3.Tab
    Case 0: ListaServ.SetFocus
    Case 1:
        Lista_custoServ.SetFocus
        txtVlrTotal_centroServ = txtValorTotal_serv
        If FunVerificaProsseguir(False, True) = False Then Exit Sub
        ProcCarregaLista_CustoServ
    Case 2:
'        Lista_empenhos_serv.SetFocus
        If FunVerificaProsseguir(False, True) = False Then Exit Sub
        Txt_qtde_total_comprada(1) = txtQtde_serv

        'Verifica se é requisição de serviço de terceiro
        If txtOrdem_serv <> "" And txtOrdem_serv <> "0" Then
            USMsgBox ("Não é permitido fazer o empenho, pois este serviço já está empenhado para uma ordem de produção."), vbExclamation, "CAPRIND v5.0"
            SSTab3.Tab = 0
            Exit Sub
        End If
        ProcCarregaListaEmpenhosServ
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab4_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    Select Case SSTab4.Tab
        Case 0: If ListaNecessidade.Visible = True Then ListaNecessidade.SetFocus
        Case 1: Lista_solicitados.SetFocus
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ID_CF_change()
On Error GoTo tratar_erro

Txt_CF = ""
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(Txt_ID_CF = "", 0, Txt_ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Txt_CF = IIf(IsNull(TBFI!IDIntClasse), "", TBFI!IDIntClasse)
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_moeda_Change()
On Error GoTo tratar_erro

If Txt_valor_moeda.Text <> "" Then
    VerifNumero = Txt_valor_moeda.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_moeda.Text = ""
        Txt_valor_moeda.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_moeda_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_valor_moeda

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_moeda_LostFocus()
On Error GoTo tratar_erro

Txt_valor_moeda = Format(Txt_valor_moeda, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAcessorias_Change()
On Error GoTo tratar_erro

If txtAcessorias.Text <> "" Then
    VerifNumero = txtFrete.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtAcessorias.Text = ""
        txtAcessorias.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor False
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAcessorias_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtAcessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAcessorias_LostFocus()
On Error GoTo tratar_erro

txtAcessorias = Format(txtAcessorias, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCFOP_prod_Change()
On Error GoTo tratar_erro

ProcCarregaCST

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

If chkAuto_serv.Value = 0 And chkManual_serv.Value = 0 Then ProcLimpaCamposServ False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesconto_Change()
On Error GoTo tratar_erro

If Chk_desc.Value = 1 Then
    If txtDesconto.Text <> "" Then
        VerifNumero = txtDesconto.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtDesconto.Text = ""
            txtDesconto.SetFocus
            Exit Sub
        End If
        valor = txtDesconto
        If valor > 100 Then
            USMsgBox ("O desconto não pode ser maior que 100."), vbExclamation, "CAPRIND v5.0"
            txtDesconto = ""
            txtDesconto.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaValor False
    ProcCalculaDesconto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto_LostFocus()
On Error GoTo tratar_erro

If txtDesconto = "" Then txtDesconto = 0
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesconto_serv_Change()
On Error GoTo tratar_erro

If Chk_desc2.Value = 1 Then
    If txtDesconto_serv.Text <> "" Then
        VerifNumero = txtDesconto_serv.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtDesconto_serv.Text = ""
            txtDesconto_serv.SetFocus
            Exit Sub
        End If
        valor = txtDesconto_serv
        If valor > 100 Then
            USMsgBox ("O desconto não pode ser maior que 100."), vbExclamation, "CAPRIND v5.0"
            txtDesconto_serv = ""
            txtDesconto_serv.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaValorServ
    ProcCalculaDescontoServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesconto_serv_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtDesconto_serv

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesconto_serv_LostFocus()
On Error GoTo tratar_erro

If txtDesconto_serv = "" Then txtDesconto_serv = 0
txtVlrDesconto_serv.Text = Format(txtVlrDesconto_serv, "###,##0.0000000000")
txtVlrUnitDesc_serv = Format(txtVlrUnitDesc_serv, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFrete_Change()
On Error GoTo tratar_erro

If txtFrete.Text <> "" Then
    VerifNumero = txtFrete.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtFrete.Text = ""
        txtFrete.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor False
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFrete_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtFrete

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFrete_LostFocus()
On Error GoTo tratar_erro

txtFrete.Text = Format(txtFrete, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIcms_Change()
On Error GoTo tratar_erro

If txtICMS.Text <> "" Then
    VerifNumero = txtICMS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtICMS.Text = ""
        txtICMS.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIcms_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtICMS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDfornecedor_Change()
On Error GoTo tratar_erro

ProcLimpafornecedor
If txtIDfornecedor <> "" Then
    VerifNumero = txtIDfornecedor
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDfornecedor = ""
        txtIDfornecedor.SetFocus
        Exit Sub
    End If
    ProcPuxafornecedor
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIPI_Change()
On Error GoTo tratar_erro

If txtIPI.Text <> "" Then
    VerifNumero = txtIPI.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIPI.Text = ""
        txtIPI.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCalculaValor(PorCF As Boolean)
On Error GoTo tratar_erro

'Zera valores
VlrIPI = 0
Qtde = 0
VlrICMS_suframa = 0

Qtde = IIf(txtQuantidade = "", 0, txtQuantidade)
If txtDesconto <> "" And txtDesconto <> "0" Then
    If txtvalorunitariodesc.Text = "" Then valor = 0 Else valor = txtvalorunitariodesc
Else
    If txtvalorunitario.Text = "" Then valor = 0 Else valor = txtvalorunitario
End If

If ChkFrete_IPI.Value = 1 Then SumTotProdutos = (valor * Qtde) + IIf(txtFrete = "", 0, txtFrete) Else SumTotProdutos = valor * Qtde

If Txt_ID_CF = "" Or Txt_ID_CF = "0" Then
    IntIPI = IIf(txtIPI = "", 0, txtIPI)
    IntICMS = IIf(txtICMS = "", 0, txtICMS)
    
    VlrIPI = (SumTotProdutos * IntIPI) / 100 'Calcula IPI
    VlrICMS_suframa = Format(((valor * Qtde) * IntICMS) / 100, "###,##0.00") 'Calcula ICMS
Else
    ProcValorImposto txtPedido, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), 0
    ProcControleImposto IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), 0
    If PorCF = True Then
        If TemIPI = "SIM" Then
            txtIPI = IntIPI
        Else
            IntIPI = 0
            txtIPI = 0
        End If
        If TemICMS = "SIM" Then
            txtICMS = IntICMS
        Else
            txtICMS = 0
            IntICMS = 0
        End If
    Else
        IntIPI = IIf(txtIPI = "", 0, txtIPI)
        IntICMS = IIf(txtICMS = "", 0, txtICMS)
    End If
    
    If SumTotProdutos = 0 Then
    SumTotProdutos = txtvlrTotal.Text
    End If
    
    VlrIPI = (SumTotProdutos * IntIPI) / 100 'Calcula IPI
    VlrIPI = (SumTotProdutos * IntIPI) / 100 'Calcula IPI
    'Calclula ICMS
    ProcCalculaBC Cmb_empresa.ItemData(Cmb_empresa.ListIndex), IIf(txtCFOP_prod = "", "0.000", txtCFOP_prod), 0, (valor * Qtde) + IIf(txtFrete = "", 0, txtFrete), VlrIPI, SomarIPI, SomarIPIST, TemReducaoBC, False, Cmb_CST_ICMS, "T", IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor
    VlrICMS_suframa = Format((BC * IntICMS) / 100, "###,##0.00")
End If

TxtvlrIpi = Format(VlrIPI, "###,##0.00")
txtvlrICMS = Format(VlrICMS_suframa, "###,##0.00")
txtvlrTotal = Format(Qtde * valor, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaValorServ()
On Error GoTo tratar_erro

If txtDesconto_serv <> "" And txtDesconto_serv <> "0" Then
    If txtVlrUnitDesc_serv.Text = "" Then valor = 0 Else valor = txtVlrUnitDesc_serv.Text
Else
    If txtValorUnit_serv.Text = "" Then valor = 0 Else valor = txtValorUnit_serv.Text
End If
If txtISSQN.Text = "" Then ValorIPI = 0 Else ValorIPI = txtISSQN.Text
If txtQtde_serv.Text = "" Then quantnovo = 0 Else quantnovo = txtQtde_serv.Text
If IsNumeric(ValorIPI) = True And IsNumeric(valor) = True Then
    txtValor_ISSQN = Format(((valor * quantnovo) * ValorIPI) / 100, "###,##0.00")
End If
If IsNumeric(quantnovo) = True And IsNumeric(valor) = True Then
    txtValorTotal_serv = Format((quantnovo * valor), "###,##0.00")
End If
valor = 0
ValorIPI = 0
quantnovo = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtipi_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtIPI

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtISSQN_Change()
On Error GoTo tratar_erro

If txtISSQN.Text <> "" Then
    VerifNumero = txtISSQN.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtISSQN.Text = ""
        txtISSQN.SetFocus
        Exit Sub
    End If
    ProcCalculaValorServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtISSQN_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtISSQN

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNomenclatura_Change()
On Error GoTo tratar_erro

If chkAuto.Value = 0 And chkManual.Value = 0 Then ProcLimpaCamposItem False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change(index As Integer)
On Error GoTo tratar_erro

If txtNreg(index) <> "" Then
    VerifNumero = txtNreg(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg(index) = ""
        txtNreg(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) <> "" Then
    VerifNumero = txtPagIr(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr(index) = ""
        txtPagIr(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOrdem_Change()
On Error GoTo tratar_erro

Cmb_OS.Clear
If txtOrdem <> "" Then
    VerifNumero = txtOrdem
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOrdem = ""
        txtOrdem.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOrdem_LostFocus()
On Error GoTo tratar_erro

With txtOrdem
    If .Text <> "" And .Text <> "0" Then
        If FunVerifOPCarregaOS(Cmb_OS, .Text, Novo_PC1, True) = False Then
            .Text = ""
            .SetFocus
        End If
    End If
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOrdem_serv_Change()
On Error GoTo tratar_erro

Cmb_OS_serv.Clear
If txtOrdem_serv <> "" Then
    VerifNumero = txtOrdem_serv
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOrdem_serv = ""
        txtOrdem_serv.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtOrdem_serv_LostFocus()
On Error GoTo tratar_erro

With txtOrdem_serv
    If .Text <> "" And .Text <> "0" Then
        If FunVerifOPCarregaOS(Cmb_OS_serv, .Text, Novo_PC2, True) = False Then
            .Text = ""
            .SetFocus
        End If
    End If
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpedido_Change()
On Error GoTo tratar_erro

If Novo_PC = True And txtPedido.Locked = True Then
VerifNPedido:
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido where Pedido = '" & txtPedido & "' and IDpedido <> " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Numero = Left(txtPedido, Len(txtPedido) - 3) + 1
        Ano = Right(Year(Date), 2)
        NumeroPedido = Numero & "/" & Ano
        txtPedido = NumeroPedido
        GoTo VerifNPedido
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentualCentro_Change()
On Error GoTo tratar_erro

If chkPercentual.Value = 1 Then
    txtValorCentro = ""
    If txtPercentualCentro.Text <> "" Then
        VerifNumero = txtPercentualCentro.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtPercentualCentro.Text = ""
            txtPercentualCentro.SetFocus
            Exit Sub
        End If
        Qtde = txtPercentualCentro
        Qtd = txtvlrTotal
        qt = (Qtd * Qtde) / 100
        txtValorCentro = Format(qt, "###,##0.00")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentualCentro_LostFocus()
On Error GoTo tratar_erro

txtPercentualCentro = Format(txtPercentualCentro, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentualCentro_Serv_Change()
On Error GoTo tratar_erro

If chkPercentual_serv.Value = 1 Then
    txtValorCentro_Serv = ""
    If txtPercentualCentro_Serv.Text <> "" Then
        VerifNumero = txtPercentualCentro_Serv.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtPercentualCentro_Serv.Text = ""
            txtPercentualCentro_Serv.SetFocus
            Exit Sub
        End If
        Qtde = txtPercentualCentro_Serv
        Qtd = txtValorTotal_serv
        qt = (Qtd * Qtde) / 100
        txtValorCentro_Serv = Format(qt, "###,##0.00")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentualCentro_Serv_LostFocus()
On Error GoTo tratar_erro


txtPercentualCentro_Serv = Format(txtPercentualCentro_Serv, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtprazo_item_LostFocus()
On Error GoTo tratar_erro

If txtprazo_item <> "__/__/____" Then
    VerifData = txtprazo_item
    ProcVerificaData
    If VerifData = False Then
        txtprazo_item = "__/__/____"
        txtprazo_item.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPrazo_serv_LostFocus()
On Error GoTo tratar_erro

If txtPrazo_serv <> "__/__/____" Then
    VerifData = txtPrazo_serv
    ProcVerificaData
    If VerifData = False Then
        txtPrazo_serv = "__/__/____"
        txtPrazo_serv.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtprazo_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtstatus_Change()
On Error GoTo tratar_erro



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_necess_Change()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_sol_Change()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (2)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_serv_Change()
On Error GoTo tratar_erro

If txtQtde_serv.Text <> "" Then
    VerifNumero = txtQtde_serv.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_serv.Text = ""
        txtQtde_serv.SetFocus
        Exit Sub
    End If
    ProcCalculaValorServ
    ProcCalculaDescontoServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_serv_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde_serv

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_serv_LostFocus()
On Error GoTo tratar_erro

txtQtde_serv.Text = Format(txtQtde_serv.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtQuantidade <> "" Then
    VerifNumero = txtQuantidade
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade = ""
        txtQuantidade.SetFocus
        Exit Sub
    End If
    ProcCalculaValor False
    ProcCalculaDesconto
    If cmbun <> Cmb_un_com Then
        txtQuantidade_est = FunFormataCasasDecimais(4, FunConversaoFinalUn(cmbun, Cmb_un_com, txtQuantidade, txtNomenclatura, True))
    Else
        txtQuantidade_est = FunFormataCasasDecimais(4, txtQuantidade)
    End If
    If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
        txtQuantidade_PC = FunCalculaQtdePC(txtNomenclatura, txtQuantidade, True, Cmb_un_com)
    Else
        txtQuantidade_PC = ""
    End If
Else
    txtQuantidade_est = ""
    txtQuantidade_PC = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQuantidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSeguro_Change()
On Error GoTo tratar_erro

If txtSeguro.Text <> "" Then
    VerifNumero = txtFrete.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtSeguro.Text = ""
        txtSeguro.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor False
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSeguro_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtSeguro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSeguro_LostFocus()
On Error GoTo tratar_erro

txtSeguro = Format(txtSeguro, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalAcessorias_Change()
On Error GoTo tratar_erro

If TxtTotalacessorias <> "" Then
    VerifNumero = TxtTotalacessorias
    ProcVerificaNumero
    If VerifNumero = False Then
        TxtTotalacessorias = ""
        TxtTotalacessorias.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalAcessorias_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus TxtTotalacessorias

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalAcessorias_LostFocus()
On Error GoTo tratar_erro
  
TxtTotalacessorias = Format(TxtTotalacessorias, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotaldesconto_Change()
On Error GoTo tratar_erro

If txtTotaldesconto <> "" Then
    VerifNumero = txtTotaldesconto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTotaldesconto = ""
        txtTotaldesconto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotaldesconto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtTotaldesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotaldesconto_LostFocus()
On Error GoTo tratar_erro
  
txtTotaldesconto = Format(txtTotaldesconto, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalFrete_Change()
On Error GoTo tratar_erro

If TxtTotalFrete <> "" Then
    VerifNumero = TxtTotalFrete
    ProcVerificaNumero
    If VerifNumero = False Then
        TxtTotalFrete = ""
        TxtTotalFrete.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalFrete_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus TxtTotalFrete

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalFrete_LostFocus()
On Error GoTo tratar_erro
  
TxtTotalFrete = Format(TxtTotalFrete, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalSeguro_Change()
On Error GoTo tratar_erro

If txtTotalSeguro <> "" Then
    VerifNumero = txtTotalSeguro
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTotalSeguro = ""
        txtTotalSeguro.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalSeguro_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtTotalSeguro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotalSeguro_LostFocus()
On Error GoTo tratar_erro
  
txtTotalSeguro = Format(txtTotalSeguro, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCentro_Change()
On Error GoTo tratar_erro

If chkValor.Value = 1 Then
    txtPercentualCentro = ""
    If txtValorCentro.Text <> "" Then
        VerifNumero = txtValorCentro.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtValorCentro.Text = ""
            txtValorCentro.SetFocus
            Exit Sub
        End If
        Qtde = txtValorCentro
        Qtd = txtvlrTotal
        qt = (Qtde * 100) / Qtd
        txtPercentualCentro = Format(qt, "###,##0.0000000000")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCentro_LostFocus()
On Error GoTo tratar_erro

txtValorCentro = Format(txtValorCentro, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCentro_Serv_Change()
On Error GoTo tratar_erro

If chkValor_serv.Value = 1 Then
    txtPercentualCentro_Serv = ""
    If txtValorCentro_Serv.Text <> "" Then
        VerifNumero = txtValorCentro_Serv.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtValorCentro_Serv.Text = ""
            txtValorCentro_Serv.SetFocus
            Exit Sub
        End If
        Qtde = txtValorCentro_Serv
        Qtd = txtValorTotal_serv
        qt = (Qtde * 100) / Qtd
        txtPercentualCentro_Serv = Format(qt, "###,##0.0000000000")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorCentro_Serv_LostFocus()
On Error GoTo tratar_erro

txtValorCentro_Serv = Format(txtValorCentro_Serv, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto_Change()
On Error GoTo tratar_erro

If Chk_valor_desc.Value = 1 Then
    If txtvalordesconto.Text <> "" Then
        VerifNumero = txtvalordesconto.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtvalordesconto.Text = ""
            txtvalordesconto.SetFocus
            Exit Sub
        End If
        valor = IIf(txtvalorunitario = "", 0, txtvalorunitario)
        Valor_Produto = txtvalordesconto
        If Valor_Produto > valor Then
            USMsgBox ("O valor do desconto não pode ser maior que o valor unitário."), vbExclamation, "CAPRIND v5.0"
            txtvalordesconto = ""
            txtvalordesconto.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaValor False
    ProcCalculaValorDesconto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalordesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto_LostFocus()
On Error GoTo tratar_erro

If txtvalordesconto = "" Then txtvalordesconto = 0
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorUnit_serv_Change()
On Error GoTo tratar_erro

If txtValorUnit_serv.Text <> "" Then
    VerifNumero = txtValorUnit_serv.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValorUnit_serv.Text = ""
        txtValorUnit_serv.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValorServ
ProcCalculaDescontoServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorUnit_serv_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtValorUnit_serv

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorUnit_serv_LostFocus()
On Error GoTo tratar_erro

txtValorUnit_serv = Format(txtValorUnit_serv, "###,##0.0000000000")
txtVlrDesconto_serv = Format(txtVlrDesconto_serv, "###,##0.0000000000")
txtVlrUnitDesc_serv = Format(txtVlrUnitDesc_serv, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitario_Change()
On Error GoTo tratar_erro

If txtvalorunitario.Text <> "" Then
    VerifNumero = txtvalorunitario.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvalorunitario.Text = ""
        txtvalorunitario.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor False
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitario_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalorunitario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitario_LostFocus()
On Error GoTo tratar_erro

txtvalorunitario = Format(txtvalorunitario, "###,##0.0000000000")
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalistapedido(Pagina As Integer)
On Error GoTo tratar_erro

If Sql_Pedido_Localizar = "" Then Exit Sub
lblRegistros(3).Caption = "Nº de registros: 0"
lblPaginas(3).Caption = "Página: 0 de: 0"
listapedido.ListItems.Clear
Set TBLISTA_Compras_Pedido = CreateObject("adodb.recordset")
'Debug.print Sql_Pedido_Localizar
TBLISTA_Compras_Pedido.Open Sql_Pedido_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Compras_Pedido.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Cotacao = 0
listapedido.ListItems.Clear
TBLISTA_Compras_Pedido.PageSize = IIf(txtNreg(3) = "", 30, txtNreg(3))
TBLISTA_Compras_Pedido.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Compras_Pedido.PageSize
ContadorReg = 1

PBLista(1).Min = 0
PBLista(1).Max = FunVerifMaxPBListaPaginacao(TBLISTA_Compras_Pedido.RecordCount - IIf(Pagina > 1, (TBLISTA_Compras_Pedido.PageSize * (Pagina - 1)), 0), TBLISTA_Compras_Pedido.PageSize)
PBLista(1).Value = 1
Contador = 0
Do While TBLISTA_Compras_Pedido.EOF = False And (ContadorReg <= TamanhoPagina)
    With listapedido.ListItems
        .Add , , TBLISTA_Compras_Pedido!IDpedido
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Compras_Pedido!Data), "", Format(TBLISTA_Compras_Pedido!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Compras_Pedido!Pedido), "", (TBLISTA_Compras_Pedido!Pedido))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Compras_Pedido!Cotacaotexto), "", TBLISTA_Compras_Pedido!Cotacaotexto)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Compras_Pedido!Fornecedor), "", (TBLISTA_Compras_Pedido!Fornecedor))
        .Item(.Count).SubItems(5) = FunCorrigeStatusPedido(TBLISTA_Compras_Pedido!Status_pedido)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Compras_Pedido!DtValidacao) = False, "Sim", "Não")
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Compras_Pedido!Data_aprovado) = False, "Sim", "Não")
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Compras_Pedido!dbl_valor_total), "00,00", Format(TBLISTA_Compras_Pedido!dbl_valor_total, "###,##0.00"))
    End With
    TBLISTA_Compras_Pedido.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista(1).Value = Contador
Loop
lblRegistros(3).Caption = "Nº de registros: " & TBLISTA_Compras_Pedido.RecordCount
If TBLISTA_Compras_Pedido.AbsolutePage = adPosBOF Then
   lblPaginas(3).Caption = "Página: 1 de: " & TBLISTA_Compras_Pedido.PageCount
ElseIf TBLISTA_Compras_Pedido.AbsolutePage = adPosEOF Then
        lblPaginas(3).Caption = "Página: " & TBLISTA_Compras_Pedido.PageCount & " de: " & TBLISTA_Compras_Pedido.PageCount
    Else
        lblPaginas(3).Caption = "Página: " & TBLISTA_Compras_Pedido.AbsolutePage - 1 & " de: " & TBLISTA_Compras_Pedido.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposItem(LimparCamposProd As Boolean)
On Error GoTo tratar_erro

If LimparCamposProd = True Then
    txtNomenclatura.Text = ""
    chkAuto.Value = 0
    chkManual.Value = 0
    txtprazo_item = "__/__/____"
    txtQuantidade.Text = ""
    txtQuantidade_PC = ""
    txtQuantidade_est = ""
    txtOrdem = ""
    Cmb_OS.ListIndex = -1
End If
txtcodproduto = 0
If txtStatus = "AGUARDANDO APROVAÇÃO" Then txtstatus_item.Text = txtStatus Else txtstatus_item.Text = "COMPRADO"
cmbReferencia.Clear
txtreferencia = ""
txtEspecificacoes.Text = ""
txtDescricao_comercial.Text = ""
txtPedidoint = ""
txtdetalheitem.Text = ""
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
txtvalorunitario.Text = ""
txtDesconto = ""
txtvalordesconto = ""
txtvalorunitariodesc = ""
txtIPI.Text = ""
TxtvlrIpi = ""
txtICMS = ""
txtvlrICMS = ""
txtvlrTotal = ""
txtObs.Text = ""
chkRemessa.Value = 0
Txt_ID_CFOP_prod = ""
txtCFOP_prod = ""
Txt_natureza_operacao_prod = ""
Txt_ID_CF = ""
Txt_CF = ""
txtFrete = ""
txtSeguro = ""
txtAcessorias = ""
ChkFrete_IPI.Value = 0
Txt_vlr_unit_ultima_compra_prod = ""
Cmb_CST_ICMS.Clear
If Novo_PC1 = True And Chk_CFOP_prod.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido_lista where IDPedido = " & IIf(txtIDPedido = "", 0, txtIDPedido) & " and Tipo = 'P' and ID_CFOP IS NOT NULL order by IDlista desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then Txt_ID_CFOP_prod = TBAbrir!ID_CFOP
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * FROM tbl_NaturezaOperacao where IDCountCfop = " & IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtCFOP_prod = IIf(IsNull(TBAbrir!ID_CFOP), "", TBAbrir!ID_CFOP)
        Txt_natureza_operacao_prod = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
    End If
    TBAbrir.Close
End If

CodigoLista1 = 0
ProcLimpaCamposCusto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposServ(LimparCamposServ As Boolean)
On Error GoTo tratar_erro

If LimparCamposServ = True Then
    txtCodigo.Text = ""
    chkAuto_serv.Value = 0
    chkManual_serv.Value = 0
    txtPrazo_serv.Text = "__/__/____"
    txtQtde_serv.Text = ""
    txtOrdem_serv = ""
    Cmb_OS_serv.ListIndex = -1
End If
txtcodproduto_serv = 0
If txtStatus = "AGUARDANDO APROVAÇÃO" Then txtStatus_serv.Text = txtStatus Else txtStatus_serv.Text = "COMPRADO"
cmbreferencia_serv.Clear
txtReferencia_serv = ""
txtDescricao_serv.Text = ""
txtDescricao_comercialServ.Text = ""
txtDetalhe_serv.Text = ""
cmbFamilia_serv.ListIndex = -1
cmbUn_serv.ListIndex = -1
Cmb_un_com_serv.ListIndex = -1
txtValorUnit_serv.Text = ""
txtDesconto_serv = ""
txtVlrDesconto_serv = ""
txtVlrUnitDesc_serv = ""
txtISSQN.Text = ""
txtValor_ISSQN = ""
txtValorTotal_serv = ""
txtObs_serv.Text = ""
Txt_ID_CFOP_serv = ""
Txt_vlr_unit_ultima_compra_serv = ""
txtCFOP_serv = ""
txtNatureza_operacao_serv = ""
If Novo_PC2 = True And Chk_CFOP_serv.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido_lista where IDpedido = " & IIf(txtIDPedido = "", 0, txtIDPedido) & " and Tipo = 'S' and ID_CFOP IS NOT NULL order by IDlista desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then Txt_ID_CFOP_serv = TBAbrir!ID_CFOP
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * FROM tbl_NaturezaOperacao where IDCountCfop = " & IIf(Txt_ID_CFOP_serv = "", 0, Txt_ID_CFOP_serv), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtCFOP_serv = IIf(IsNull(TBAbrir!ID_CFOP), "", TBAbrir!ID_CFOP)
        txtNatureza_operacao_serv = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
    End If
    TBAbrir.Close
End If

CodigoLista3 = 0
ProcLimpaCamposCustoServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirLista()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Listprod
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select * from compras_pedido_lista where idlista = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then
                If TBCompras_Lista!ID_Requisicao <> 0 And TBCompras_Lista!ID_cotacao = 0 Then
                    Conexao.Execute "Update Compras_pedido_lista_custo Set IDpedido = 0, Valor = 0 WHERE IDLista = " & .ListItems(InitFor)
                    TBCompras_Lista!IDpedido = 0
                    TBCompras_Lista!Quant_Comp = 0
                    TBCompras_Lista!preco_unitario = 0
                    TBCompras_Lista!IPI = 0
                    TBCompras_Lista!preco_total = 0
                    TBCompras_Lista!Status_Item = "REQUISIT."
                    TBCompras_Lista!vlrICMS = 0
                    TBCompras_Lista!VlrIPI = 0
                    TBCompras_Lista!ICMS = 0
                    TBCompras_Lista.Update
                Else
                    If TBCompras_Lista!ID_cotacao <> 0 Then
                        Conexao.Execute "Update Compras_pedido_lista_custo Set IDpedido = 0, Valor = 0 WHERE IDLista = " & .ListItems(InitFor)
                        Conexao.Execute "Update CF set CF.IDPedido = 0, CF.aprovadoforn = 'False', CF.naprovadoforn = 'True' from (Cotacao_fornecedor CF INNER JOIN Cotacao_item CI ON CF.IDitem = CI.ID) INNER JOIN Compras_pedido_lista CPL ON CPL.IDlista = CI.iditemlista where CPL.IDLista = " & .ListItems(InitFor)
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Cotacao_fornecedor where IDcot = " & TBCompras_Lista!ID_cotacao & " and IDPedido <> 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = True Then
                            Conexao.Execute "Update Compras_Cotacao Set dataaprovada = NULL, statuscotacao = 'LIBERADA' where ID_cotacao = " & TBCompras_Lista!ID_cotacao
                        End If
                        TBAbrir.Close
                        TBCompras_Lista!IDpedido = 0
                        TBCompras_Lista!Quant_Comp = 0
                        TBCompras_Lista!preco_unitario = 0
                        TBCompras_Lista!IPI = 0
                        TBCompras_Lista!preco_total = 0
                        TBCompras_Lista!Status_Item = "COTANDO"
                        TBCompras_Lista!vlrICMS = 0
                        TBCompras_Lista!VlrIPI = 0
                        TBCompras_Lista!ICMS = 0
                        TBCompras_Lista.Update
                    Else
                        Conexao.Execute "DELETE from compras_pedido_lista WHERE IDLista = " & .ListItems(InitFor)
                        Conexao.Execute "DELETE from Compras_pedido_lista_custo WHERE IDLista = " & .ListItems(InitFor)
                        Conexao.Execute "DELETE from vendas_carteira_alteracoes where ID_carteira = " & .ListItems(InitFor) & " and Tipo = 'CPE'"
                        Conexao.Execute "DELETE from Compras_pedido_lista_empenhos WHERE IDLista = " & .ListItems(InitFor)
                    End If
                End If
            End If
            TBCompras_Lista.Close

            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Excluir produto"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº pedido: " & txtPedido
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
            
            'Excluir fornecedor do produto
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select codproduto from projproduto where desenho = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select CPL.* from compras_pedido_lista CPL INNER JOIN compras_pedido CP on CPL.IDpedido = CP.IDpedido where CPL.codproduto = " & TBItem!Codproduto & " and CP.IDfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = True Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select NF.*, NFP.codproduto from tbl_detalhes_nota NFP INNER JOIN tbl_dados_nota_fiscal NF ON NFP.id_nota = NF.ID where NFP.int_Cod_Produto = '" & .ListItems(InitFor).SubItems(1) & "' and NF.Id_Int_Cliente = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then Conexao.Execute "DELETE from Projproduto_fornecedor WHERE codproduto = " & .ListItems(InitFor) & " and IDfornecedor = " & txtIDfornecedor
                End If
                TBOrdem.Close
            End If
            TBItem.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    txtStatus = FunAtualizaStatusPC(txtIDPedido)
    TXTIDLista = 0
    ProcLimpaCamposItem True
    Frame1(12).Enabled = False
    ProcAtualizalista
    Novo_PC1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirServ()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaServ
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select * from compras_pedido_lista where idlista = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then
                If TBCompras_Lista!ID_Requisicao <> 0 And TBCompras_Lista!ID_cotacao = 0 Then
                    Conexao.Execute "Update Compras_pedido_lista_custo Set IDpedido = 0, Valor = 0 WHERE IDLista = " & .ListItems(InitFor)
                    TBCompras_Lista!IDpedido = 0
                    TBCompras_Lista!Quant_Comp = 0
                    TBCompras_Lista!preco_unitario = 0
                    TBCompras_Lista!IPI = 0
                    TBCompras_Lista!preco_total = 0
                    TBCompras_Lista!Status_Item = "REQUISIT."
                    TBCompras_Lista!vlrICMS = 0
                    TBCompras_Lista!VlrIPI = 0
                    TBCompras_Lista!VlrISSQN = 0
                    TBCompras_Lista!ICMS = 0
                    TBCompras_Lista.Update
                Else
                    If TBCompras_Lista!ID_cotacao <> 0 Then
                        Conexao.Execute "Update Compras_pedido_lista_custo Set IDpedido = 0, Valor = 0 WHERE IDLista = " & .ListItems(InitFor)
                        Conexao.Execute "Update CF set CF.IDPedido = 0, CF.aprovadoforn = 'False', CF.naprovadoforn = 'True' from (Cotacao_fornecedor CF INNER JOIN Cotacao_item CI ON CF.IDitem = CI.ID) INNER JOIN Compras_pedido_lista CPL ON CPL.IDlista = CI.iditemlista where CPL.IDLista = " & .ListItems(InitFor)
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Cotacao_fornecedor where IDcot = " & TBCompras_Lista!ID_cotacao & " and IDPedido <> 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = True Then
                            Conexao.Execute "Update Compras_Cotacao Set dataaprovada = NULL, statuscotacao = 'LIBERADA' where ID_cotacao = " & TBCompras_Lista!ID_cotacao
                        End If
                        TBAbrir.Close
                        TBCompras_Lista!IDpedido = 0
                        TBCompras_Lista!Quant_Comp = 0
                        TBCompras_Lista!preco_unitario = 0
                        TBCompras_Lista!IPI = 0
                        TBCompras_Lista!preco_total = 0
                        TBCompras_Lista!Status_Item = "COTANDO"
                        TBCompras_Lista!vlrICMS = 0
                        TBCompras_Lista!VlrIPI = 0
                        TBCompras_Lista!VlrISSQN = 0
                        TBCompras_Lista!ICMS = 0
                        TBCompras_Lista.Update
                    Else
                        Conexao.Execute "DELETE from compras_pedido_lista WHERE IDLista = " & .ListItems(InitFor)
                        Conexao.Execute "DELETE from Compras_pedido_lista_custo WHERE IDLista = " & .ListItems(InitFor)
                        Conexao.Execute "DELETE from vendas_carteira_alteracoes where ID_carteira = " & .ListItems(InitFor) & " and Tipo = 'CPE'"
                        Conexao.Execute "DELETE from Compras_pedido_lista_empenhos WHERE IDLista = " & .ListItems(InitFor)
                    End If
                End If
            End If
            TBCompras_Lista.Close

            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Excluir serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº pedido: " & txtCotacao & " - Rev.: " & txtrevisao
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
                        
            'Excluir fornecedor do serviço
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select codproduto from projproduto where desenho = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select CPL.* from compras_pedido_lista CPL INNER JOIN compras_pedido CP on CPL.IDpedido = CP.IDpedido where CPL.codproduto = " & TBItem!Codproduto & " and CP.IDfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = True Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select NF.*, NFP.codproduto from tbl_detalhes_nota NFP INNER JOIN tbl_dados_nota_fiscal NF on NFP.id_nota = NF.ID where NFP.int_Cod_Produto = '" & txtCodigo & "' and NF.Id_Int_Cliente = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then Conexao.Execute "DELETE from Projproduto_fornecedor WHERE codproduto = " & txtcodproduto_serv & " and IDfornecedor = " & txtIDfornecedor
                End If
                TBOrdem.Close
            End If
            TBItem.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    txtStatus = FunAtualizaStatusPC(txtIDPedido)
    txtIDLista_serv = 0
    ProcLimpaCamposServ True
    Frame1(7).Enabled = False
    ProcAtualizalistaServ
    Novo_PC2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpafornecedor()
On Error GoTo tratar_erro

txtFornecedor.Text = ""
txtEmail = ""
txttelefone = ""
txtContato = ""
txtendereco = ""
txtTipo_bairro = ""
txtTipo_endereco = ""
txtNumero = ""
txtCidade = ""
txtBairro = ""
txtuf = ""
txtFax = ""
txtCategoria.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaDesconto()
On Error GoTo tratar_erro

If txtvalorunitario.Text <> "" And txtDesconto <> "" Then
    If IsNumeric(txtvalorunitario.Text) = True Then
        a = Format(txtvalorunitario.Text, "###,##0.0000000000")
        c = IIf(txtDesconto = "", 0, txtDesconto)
        D = (a * c) / 100
        txtvalordesconto.Text = Format(D, "###,##0.0000000000")
        txtvalorunitariodesc.Text = Format(a - D, "###,##0.0000000000")
        ProcCalculaValores
    End If
Else
    txtvalordesconto = "0,00000"
    txtvalorunitariodesc = IIf(txtvalorunitario = "", "0,00000", txtvalorunitario)
    'TxtVlrTotal = "0,00"
    'TxtvlrIpi = "0,00"
    'TxtVlrIcms = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaDescontoServ()
On Error GoTo tratar_erro

If txtValorUnit_serv.Text <> "" And txtDesconto_serv <> "" Then
    If IsNumeric(txtValorUnit_serv.Text) = True Then
        a = Format(txtValorUnit_serv.Text, "###,##0.0000000000")
        c = IIf(txtDesconto_serv = "", 0, txtDesconto_serv)
        D = (a * c) / 100
        txtVlrDesconto_serv.Text = Format(D, "###,##0.0000000000")
        txtVlrUnitDesc_serv.Text = Format(a - D, "###,##0.0000000000")
        ProcCalculaValoresServ
    End If
Else
    txtVlrDesconto_serv = "0,00000"
    txtVlrUnitDesc_serv = IIf(txtValorUnit_serv = "", "0,00000", txtValorUnit_serv)
    txtValorTotal_serv = "0,00"
    txtValor_ISSQN = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorDesconto()
On Error GoTo tratar_erro

If txtvalorunitario.Text <> "" Then
    If IsNumeric(txtvalorunitario.Text) = True Then
        quantestoque = txtvalorunitario.Text
        QuantSolicitado = IIf(txtvalordesconto = "", 0, txtvalordesconto)
        If quantestoque <> 0 Then QuantEmpenho = (QuantSolicitado * 100) / quantestoque Else QuantEmpenho = 0
        txtDesconto.Text = QuantEmpenho
        txtvalorunitariodesc.Text = Format(quantestoque - QuantSolicitado, "###,##0.0000000000")
    Else
        Exit Sub
    End If
    ProcCalculaValores
Else
    txtvalordesconto = "0,00000"
    txtvalorunitariodesc = IIf(txtvalorunitario = "", "0,00000", txtvalorunitario)
    txtvlrTotal = "0,00"
    TxtvlrIpi = "0,00"
    txtvlrICMS = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorDescontoServ()
On Error GoTo tratar_erro

If txtValorUnit_serv.Text <> "" Then
    If IsNumeric(txtValorUnit_serv.Text) = True Then
        quantestoque = txtValorUnit_serv.Text
        QuantSolicitado = IIf(txtVlrDesconto_serv = "", 0, txtVlrDesconto_serv)
        If quantestoque <> 0 Then QuantEmpenho = (QuantSolicitado * 100) / quantestoque Else QuantEmpenho = 0
        txtDesconto_serv.Text = QuantEmpenho
        txtVlrUnitDesc_serv.Text = Format(quantestoque - QuantSolicitado, "###,##0.0000000000")
    Else
        Exit Sub
    End If
    ProcCalculaValoresServ
Else
    txtVlrDesconto_serv = "0,00000"
    txtVlrUnitDesc_serv = IIf(txtValorUnit_serv = "", "0,00000", txtValorUnit_serv)
    txtValorTotal_serv = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValores()
On Error GoTo tratar_erro

If txtQuantidade.Text = "" Or txtvalorunitario.Text = "" Then Exit Sub
'Atribui valores
If txtvalorunitariodesc.Text = "" Or txtvalorunitariodesc.Text = "0,00000" Then
    txtvlrTotal = Format(txtvalorunitario.Text * txtQuantidade, "###,##0.00")
Else
    txtvlrTotal = Format(txtvalorunitariodesc.Text * txtQuantidade, "###,##0.00")
End If
ProcCalculaValor False
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValoresServ()
On Error GoTo tratar_erro

If txtQtde_serv.Text = "" Or txtValorUnit_serv.Text = "" Then Exit Sub
'Atribui valores
If txtVlrUnitDesc_serv.Text = "" Or txtVlrUnitDesc_serv.Text = "0,00000" Then
    txtValorTotal_serv = Format(txtValorUnit_serv.Text * txtQtde_serv, "###,##0.00")
Else
    txtValorTotal_serv = Format(txtVlrUnitDesc_serv * txtQtde_serv, "###,##0.00")
End If
ProcCalculaValorServ
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaTotais()
0 On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select * from compras_pedido where idpedido = " & txtIDPedido.Text
'Debug.print StrSql

TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    txt_vlrtotalprod.Text = IIf(IsNull(TBFIltro!dbl_Valor_Total_Produtos), "0,00", Format(TBFIltro!dbl_Valor_Total_Produtos, "###,##0.00"))
    txttotalservicos.Text = IIf(IsNull(TBFIltro!dbl_valor_total_servicos), "0,00", Format(TBFIltro!dbl_valor_total_servicos, "###,##0.00"))
    txtTotaldesconto.Text = IIf(IsNull(TBFIltro!TotalDesconto), "0,00", Format(TBFIltro!TotalDesconto, "###,##0.00"))
    txt_TotalIPI.Text = IIf(IsNull(TBFIltro!dbl_Valor_Total_IPI), "0,00", Format(TBFIltro!dbl_Valor_Total_IPI, "###,##0.00"))
    txt_ICMS_ST.Text = IIf(IsNull(TBFIltro!dbl_Valor_ICMS_Subst), "0,00", Format(TBFIltro!dbl_Valor_ICMS_Subst, "###,##0.00"))
    txt_BaseICMS.Text = IIf(IsNull(TBFIltro!dbl_Base_ICMS), "0,00", Format(TBFIltro!dbl_Base_ICMS, "###,##0.00"))
    txt_vlrICMS.Text = IIf(IsNull(TBFIltro!dbl_Valor_ICMS), "0,00", Format(TBFIltro!dbl_Valor_ICMS, "###,##0.00"))
    txt_baseICMS_ST.Text = IIf(IsNull(TBFIltro!dbl_Base_ICMS_Subst), "0,00", Format(TBFIltro!dbl_Base_ICMS_Subst, "###,##0.00"))
    TxtTotalFrete = IIf(IsNull(TBFIltro!Total_Frete), "0,00", Format(TBFIltro!Total_Frete, "###,##0.00"))
    txtTotalSeguro = IIf(IsNull(TBFIltro!Total_Seguro), "0,00", Format(TBFIltro!Total_Seguro, "###,##0.00"))
    TxtTotalacessorias = IIf(IsNull(TBFIltro!Total_Acessorias), "0,00", Format(TBFIltro!Total_Acessorias, "###,##0.00"))
    txtTotalPedido.Text = IIf(IsNull(TBFIltro!dbl_valor_total), "0,00", Format(TBFIltro!dbl_valor_total, "###,##0.00"))
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiarPedido()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from compras_pedido where idpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtPedido = FunCriarNovoNumero
    idpedido_compra = txtIDPedido
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from compras_pedido", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!Pedido = txtPedido
    TBGravar!ID_empresa = TBAbrir!ID_empresa
    TBGravar!Responsavel = pubUsuario
    TBGravar!Data = Date
    TBGravar!Status_pedido = "AGUARDANDO APROVAÇÃO"
    TBGravar!IDFornecedor = TBAbrir!IDFornecedor
    TBGravar!Fornecedor = TBAbrir!Fornecedor
    TBGravar!CPF_CNPJ = TBAbrir!CPF_CNPJ
    TBGravar!Categoria = TBAbrir!Categoria
    TBGravar!contato = TBAbrir!contato
    TBGravar!Tipo_endereco = TBAbrir!Tipo_endereco
    TBGravar!Endereco = TBAbrir!Endereco
    TBGravar!Numero = TBAbrir!Numero
    TBGravar!Tipo_bairro = TBAbrir!Tipo_bairro
    TBGravar!Bairro = TBAbrir!Bairro
    TBGravar!Cidade = TBAbrir!Cidade
    TBGravar!Estado = TBAbrir!Estado
    TBGravar!Email = TBAbrir!Email
    TBGravar!fone = TBAbrir!fone
    TBGravar!Fax = TBAbrir!Fax
    TBGravar!N_referencia = TBAbrir!N_referencia
    TBGravar!Descricao_referencia = TBAbrir!Descricao_referencia
    TBGravar!dbl_valor_total = TBAbrir!dbl_valor_total
    TBGravar!dbl_Valor_Total_IPI = TBAbrir!dbl_Valor_Total_IPI
    TBGravar!dbl_Valor_Total_Produtos = TBAbrir!dbl_Valor_Total_Produtos
    TBGravar!dbl_Valor_ICMS = TBAbrir!dbl_Valor_ICMS
    TBGravar!dbl_Base_ICMS = TBAbrir!dbl_Base_ICMS
    TBGravar!dbl_valor_total_servicos = TBAbrir!dbl_valor_total_servicos
    TBGravar!TotalDesconto = TBAbrir!TotalDesconto
    TBGravar!SubTotal = TBAbrir!SubTotal
    TBGravar!dbl_Base_ICMS_Subst = TBAbrir!dbl_Base_ICMS_Subst
    TBGravar!dbl_Valor_ICMS_Subst = TBAbrir!dbl_Valor_ICMS_Subst
    TBGravar!Total_Frete = TBAbrir!Total_Frete
    TBGravar!Total_Seguro = TBAbrir!Total_Seguro
    TBGravar!Total_Acessorias = TBAbrir!Total_Acessorias
    TBGravar.Update
    txtIDPedido = TBGravar!IDpedido
    TBGravar.Close
    ProcCopiarComercial
    ProcCopiarProdutoServ
    USMsgBox ("Pedido de compra copiado com sucesso."), vbInformation, "CAPRIND v5.0"
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiarComercial()
On Error GoTo tratar_erro

Set TBNivel1 = CreateObject("adodb.recordset")
TBNivel1.Open "Select * FROM compras_comercial WHERE IDpedido = " & idpedido_compra, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel1.EOF = False Then
    Set TBNivel2 = CreateObject("adodb.recordset")
    TBNivel2.Open "Select * FROM compras_comercial", Conexao, adOpenKeyset, adLockOptimistic
    TBNivel2.AddNew
    TBNivel2!condicoes = TBNivel1!condicoes
    TBNivel2!IDpedido = txtIDPedido.Text
    TBNivel2!Embalagem = TBNivel1!Embalagem
    TBNivel2!ID_entrega = TBNivel1!ID_entrega
    TBNivel2!localentrega = TBNivel1!localentrega
    TBNivel2!Tipo_transp = TBNivel1!Tipo_transp
    TBNivel2!Idtransporte = TBNivel1!Idtransporte
    TBNivel2!Observacoes = TBNivel1!Observacoes
    TBNivel2!Prazo = TBNivel1!Prazo
    TBNivel2!Banco = TBNivel1!Banco
    TBNivel2!Agencia = TBNivel1!Agencia
    TBNivel2!Conta = TBNivel1!Conta
    TBNivel2!Escopo = TBNivel1!Escopo
    TBNivel2!Moeda = TBNivel1!Moeda
    TBNivel2!Valor_moeda = TBNivel1!Valor_moeda
    TBNivel2.Update
    TBNivel2.Close
End If
TBNivel1.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiarProdutoServ()
On Error GoTo tratar_erro

Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from compras_pedido_lista where idpedido = " & idpedido_compra, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    Do While TBCompras_Lista.EOF = False
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
        TBProduto.AddNew
        TBProduto!IDpedido = txtIDPedido
        TBProduto!Status_Item = "AGUARDANDO APROVAÇÃO"
        TBProduto!Codproduto = TBCompras_Lista!Codproduto
        TBProduto!Desenho = TBCompras_Lista!Desenho
        TBProduto!Descricao = TBCompras_Lista!Descricao
        TBProduto!Descricao_comercial = TBCompras_Lista!Descricao_comercial
        TBProduto!detalheitem = TBCompras_Lista!detalheitem
        TBProduto!Quant_Comp = TBCompras_Lista!Quant_Comp
        TBProduto!Quant_Comp_PC = TBCompras_Lista!Quant_Comp_PC
        TBProduto!Desconto = TBCompras_Lista!Desconto
        TBProduto!ValorDesconto = TBCompras_Lista!ValorDesconto
        TBProduto!preco_unitario_desconto = TBCompras_Lista!preco_unitario_desconto
        TBProduto!preco_unitario = TBCompras_Lista!preco_unitario
        TBProduto!preco_total = TBCompras_Lista!preco_total
        TBProduto!IPI = TBCompras_Lista!IPI
        TBProduto!Familia = TBCompras_Lista!Familia
        TBProduto!Un = TBCompras_Lista!Un
        TBProduto!Unidade_com = TBCompras_Lista!Unidade_com
        TBProduto!BC_ICMS = TBCompras_Lista!BC_ICMS
        TBProduto!ICMS = TBCompras_Lista!ICMS
        TBProduto!vlrICMS = TBCompras_Lista!vlrICMS
        TBProduto!VlrIPI = TBCompras_Lista!VlrIPI
        TBProduto!Prazo = TBCompras_Lista!Prazo
        TBProduto!Obs_pedido = TBCompras_Lista!Obs_pedido
        TBProduto!Remessa = TBCompras_Lista!Remessa
        TBProduto!Tipo = TBCompras_Lista!Tipo
        TBProduto!ISSQN = TBCompras_Lista!ISSQN
        TBProduto!VlrISSQN = TBCompras_Lista!VlrISSQN
        TBProduto!ID_CFOP = TBCompras_Lista!ID_CFOP
        TBProduto!ID_CF = TBCompras_Lista!ID_CF
        TBProduto!CST = TBCompras_Lista!CST
        TBProduto!Valor_ICMS_ST = TBCompras_Lista!Valor_ICMS_ST
        TBProduto!BC_ICMS_ST = TBCompras_Lista!BC_ICMS_ST
        TBProduto!Frete = TBCompras_Lista!Frete
        TBProduto!Seguro = TBCompras_Lista!Seguro
        TBProduto!Acessorias = TBCompras_Lista!Acessorias
        TBProduto!Acessorias = TBCompras_Lista!Frete_IPI
        TBProduto!Qtde_estoque = TBCompras_Lista!Qtde_estoque
        TBProduto.Update
        
        'Copiar centro de custo
        Set TBCarteira = CreateObject("adodb.recordset")
        TBCarteira.Open "Select * from Compras_pedido_lista_custo where IDLista = " & TBCompras_Lista!IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            Do While TBCarteira.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from Compras_pedido_lista_custo", Conexao, adOpenKeyset, adLockOptimistic
                TBGravar.AddNew
                TBGravar!IDpedido = txtIDPedido
                TBGravar!IDlista = TBProduto!IDlista
                TBGravar!ID_CC = TBCarteira!ID_CC
                TBGravar!valor = TBCarteira!valor
                TBGravar!Percentual = TBCarteira!Percentual
                TBGravar!Data = Date
                TBGravar!Responsavel = pubUsuario
                TBGravar.Update
                TBCarteira.MoveNext
            Loop
        End If
        TBCarteira.Close
        
        TBProduto.Close
        TBCompras_Lista.MoveNext
    Loop
End If
TBCompras_Lista.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEscopoForn()
On Error GoTo tratar_erro

txtEscopo = ""
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM compras_comercial WHERE IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtEscopo = IIf(IsNull(TBProduto!Escopo), "", TBProduto!Escopo)
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoAuto()
On Error GoTo tratar_erro

txtNomenclatura = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtNomenclatura, txtreferencia, 0, txtEspecificacoes, txtEspecificacoes, cmbfamilia, 0, 0, txtvalorunitario, cmbun, Cmb_un_com, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), True, False, True, False, 0, "P", "", 0, 0, 0, "", txtIDfornecedor, txtFornecedor, "F")
txtcodproduto = Codproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcnovoServicoAuto()
On Error GoTo tratar_erro

txtCodigo = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtCodigo, txtReferencia_serv, 0, txtDescricao_serv, txtDescricao_serv, cmbFamilia_serv, 0, 0, txtValorUnit_serv, cmbUn_serv, Cmb_un_com_serv, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", txtIDfornecedor, txtFornecedor, "F")
txtcodproduto_serv = Codproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoManual()
On Error GoTo tratar_erro

txtNomenclatura = FunCriaNovoProdServ(True, "", txtNomenclatura, txtreferencia, 0, txtEspecificacoes, txtEspecificacoes, cmbfamilia, 0, 0, txtvalorunitario, cmbun, Cmb_un_com, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), True, False, True, False, 0, "P", "", 0, 0, 0, "", txtIDfornecedor, txtFornecedor, "F")
txtcodproduto = Codproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoServicoManual()
On Error GoTo tratar_erro

txtCodigo = FunCriaNovoProdServ(True, "", txtCodigo, txtReferencia_serv, 0, txtDescricao_serv, txtDescricao_serv, cmbFamilia_serv, 0, 0, txtValorUnit_serv, cmbUn_serv, Cmb_un_com_serv, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", txtIDfornecedor, txtFornecedor, "F")
txtcodproduto_serv = Codproduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxafornecedor()
On Error GoTo tratar_erro

If Novo_PC = True Then TextoFiltro = "and DtValidacao IS NOT NULL and status <> 'Bloqueado' and Prospecto = 'False'" Else TextoFiltro = ""
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_fornecedores where idcliente = " & IIf(txtIDfornecedor = "", 0, txtIDfornecedor) & " " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIDfornecedor = TBAbrir!IDCliente
    txtFornecedor = IIf(IsNull(TBAbrir!Nome_Razao), "", TBAbrir!Nome_Razao)
    txtEmail = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
    txttelefone = IIf(IsNull(TBAbrir!Telefones), "", TBAbrir!Telefones)
    txtTipo_endereco = IIf(IsNull(TBAbrir!Tipo_endereco), "", TBAbrir!Tipo_endereco)
    txtendereco = IIf(IsNull(TBAbrir!Endereco), "", Trim(TBAbrir!Endereco))
    txtNumero = IIf(IsNull(TBAbrir!Numero), "", Trim(TBAbrir!Numero))
    txtCidade = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
    txtTipo_bairro = IIf(IsNull(TBAbrir!Tipo_bairro), "", TBAbrir!Tipo_bairro)
    txtBairro = IIf(IsNull(TBAbrir!Bairro), "", TBAbrir!Bairro)
    txtuf = IIf(IsNull(TBAbrir!Estado), "", TBAbrir!Estado)
    txtFax = IIf(IsNull(TBAbrir!Fax), "", TBAbrir!Fax)
    txtCategoria.Text = IIf(IsNull(TBAbrir!Categoria), "", TBAbrir!Categoria)
End If
TBAbrir.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlrDesconto_serv_Change()
On Error GoTo tratar_erro

If Chk_valor_desc2.Value = 1 Then
    If txtVlrDesconto_serv.Text <> "" Then
        VerifNumero = txtVlrDesconto_serv.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtVlrDesconto_serv.Text = ""
            txtVlrDesconto_serv.SetFocus
            Exit Sub
        End If
        valor = IIf(txtValorUnit_serv = "", 0, txtValorUnit_serv)
        Valor_Produto = txtVlrDesconto_serv
        If Valor_Produto > valor Then
            USMsgBox ("O valor do desconto não pode ser maior que o valor unitário."), vbExclamation, "CAPRIND v5.0"
            txtVlrDesconto_serv = ""
            txtVlrDesconto_serv.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaValorServ
    ProcCalculaValorDescontoServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlrDesconto_serv_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtVlrDesconto_serv

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtVlrDesconto_serv_LostFocus()
On Error GoTo tratar_erro

If txtVlrDesconto_serv = "" Then txtVlrDesconto_serv = 0
txtVlrDesconto_serv = Format(txtVlrDesconto_serv, "###,##0.0000000000")
txtVlrUnitDesc_serv = Format(txtVlrUnitDesc_serv, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarItem_custo()
On Error GoTo tratar_erro

If Frame1(13).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_centro = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_centro.SetFocus
    Exit Sub
End If
If chkPercentual.Value = 1 Then
    Valor1 = IIf(txtPercentualCentro = "", 0, txtPercentualCentro)
    If Valor1 = 0 Then
        NomeCampo = "o percentual"
        ProcVerificaAcao
        txtPercentualCentro.SetFocus
        Exit Sub
    End If
End If
Valor1 = IIf(txtValorCentro = "", 0, txtValorCentro)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    chkValor.Value = 1
    txtValorCentro.SetFocus
    Exit Sub
End If

'Verifica se o valor do centro de custo já ultrapassou o valor do item
Qtde = 0
Qtd = IIf(txtValorCentro = "", 0, txtValorCentro)
qt = txtvlrTotal
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(valor) as qtde from Compras_pedido_lista_custo where IDLista = " & TXTIDLista & " and ID <> " & txtIDCentro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
End If
TBAbrir.Close
If Format((Qtde + Qtd), "###,##0.00") > qt Then
    USMsgBox ("Não é permitido salvar, pois o valor do centro de custo ultrapassou o valor total do produto."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
ID_CC = 0
'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If Novo_PC1_Custo = False Then ID_CC = Lista_custo.SelectedItem.ListSubItems(5)

    If ID_CC <> Cmb_centro.ItemData(Cmb_centro.ListIndex) Then
        Formulario = "Compras/Autorização de centro de custo sem previsão"
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select ID_PC from projproduto where desenho = '" & txtNomenclatura & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            If USMsgBox("O produto não possui conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
        Else
            Set TBCQ = CreateObject("adodb.recordset")
            TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & Cmb_centro.ItemData(Cmb_centro.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
            If TBCQ.EOF = True Then
                If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
            Else
                Permitido = True
            End If
            TBCQ.Close
        End If
        TBProduto.Close
        If Permitido = False Then Exit Sub
    End If
End If
TBTempo.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_pedido_lista where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Novo_PC1_Custo = True Then TextoFiltro = "ID_CC = " & Cmb_centro.ItemData(Cmb_centro.ListIndex) Else TextoFiltro = "ID_CC = " & Lista_custo.SelectedItem.ListSubItems(5)
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from Compras_pedido_lista_custo where IDlista = " & TBAbrir!IDlista & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar!IDpedido = txtIDPedido
        TBGravar!IDlista = TBAbrir!IDlista
        TBGravar!Responsavel = pubUsuario
        TBGravar!Data = Date
        Evento = "Novo centro de custo do produto"
    Else
        If txtResponsavel_aprovacao <> "" Then
            If txtResponsavel_aprovacao <> pubUsuario Then
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "alterar", True, True) = False Then Exit Sub
            End If
        Else
            If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "alterar", True, True) = False Then Exit Sub
        End If
        If FunVerifSatus("alterar este centro de custo", True) = False Then Exit Sub
        If FunVerifSatusProdServ(txtstatus_item, "alterar este centro de custo", True, True) = False Then Exit Sub
        
        Evento = "Alterar centro de custo do produto"
    End If
    TBGravar!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex)
    TBGravar!Percentual = txtPercentualCentro
    
    Qtde = txtPercentualCentro
    Qtd = TBAbrir!preco_total
    qt = (Qtd * Qtde) / 100
    TBGravar!valor = Format(qt, "###,##0.00")
    
    TBGravar.Update
    TBGravar.Close
    
    '==================================
    Modulo = "Compras/Pedido"
    ID_documento = TBAbrir!IDlista
    Documento = "Nº pedido: " & txtPedido
    Documento1 = "Cód. interno: " & TBAbrir!Desenho & " - Centro de custo: " & Cmb_centro
    ProcGravaEvento
    '==================================
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID from Compras_pedido_lista_custo where IDlista = " & TXTIDLista & " and ID_CC = " & Cmb_centro.ItemData(Cmb_centro.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtIDCentro = TBAbrir!ID
    End If
    TBAbrir.Close
    
    ProcCarregaLista_Custo
    If Novo_PC1_Custo = True Then
        USMsgBox ("Novo centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        If Lista_custo.ListItems.Count <> 0 And CodigoLista2 <> 0 Then
            Lista_custo.SelectedItem = Lista_custo.ListItems(CodigoLista2)
            Lista_custo.SetFocus
        End If
    End If
    Novo_PC1_Custo = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirLista_Custo()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_custo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) centro(s) de custo do produto " & txtNomenclatura & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select * from Compras_pedido_lista_custo where ID = " & .ListItems(InitFor) & " and ID_requisicao <> 0 and ID_requisicao is not null", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then
                TBCompras_Lista!IDpedido = 0
                TBCompras_Lista!valor = 0
                TBCompras_Lista!Percentual = 0
                TBCompras_Lista.Update
            Else
                Conexao.Execute "DELETE from Compras_pedido_lista_custo WHERE ID = " & .ListItems(InitFor)
            End If
            TBCompras_Lista.Close
            
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Excluir centro de custo do produto"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº pedido: " & txtPedido
            If .ListItems(InitFor).SubItems(1) <> "" Then CC = .ListItems(InitFor).SubItems(1) & " - " & .ListItems(InitFor).SubItems(2) Else CC = .ListItems(InitFor).SubItems(2)
            Documento1 = "Cód. interno: " & txtNomenclatura & " - Centro de custo: " & CC
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centro(s) de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Centro(s) de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCusto
    Frame1(13).Enabled = False
    ProcCarregaLista_Custo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpenhoProd()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With Lista_empenhos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where ID = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Excluir empenho do produto"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº pedido: " & txtPedido & " - Cód. interno: " & txtNomenclatura
            Documento1 = "Pedido int.: " & .ListItems(InitFor).ListSubItems(2) & " - Rev.: " & .ListItems(InitFor).ListSubItems(3) & " - Cód. interno: " & .ListItems(InitFor).ListSubItems(5) & " - Rev.: " & .ListItems(InitFor).ListSubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaEmpenhosProd
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpenhoServ()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With Lista_empenhos_serv
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where ID = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Excluir empenho do serviço"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº pedido: " & txtPedido & " - Cód. interno: " & txtCodigo
            Documento1 = "Pedido int.: " & .ListItems(InitFor).ListSubItems(2) & " - Rev.: " & .ListItems(InitFor).ListSubItems(3) & " - Cód. interno: " & .ListItems(InitFor).ListSubItems(5) & " - Rev.: " & .ListItems(InitFor).ListSubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaEmpenhosServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposCusto()
On Error GoTo tratar_erro

txtIDCentro = 0
ProcCarregaComboSetor Cmb_centro, "Setor IS NOT NULL and DtBloq IS NULL and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
txtValorCentro = ""
chkValor.Value = 0
chkPercentual.Value = 0
txtValorCentro.Locked = True
txtValorCentro.TabStop = False
With txtPercentualCentro
    .Text = ""
    .Locked = True
    .TabStop = False
End With
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarServ_custo()
On Error GoTo tratar_erro

If Frame1(8).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_centro_servico = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_centro_servico.SetFocus
    Exit Sub
End If
If chkPercentual_serv.Value = 1 Then
    Valor1 = IIf(txtPercentualCentro_Serv = "", 0, txtPercentualCentro_Serv)
    If Valor1 = 0 Then
        NomeCampo = "o percentual"
        ProcVerificaAcao
        txtPercentualCentro_Serv.SetFocus
        Exit Sub
    End If
End If
Valor1 = IIf(txtValorCentro_Serv = "", 0, txtValorCentro_Serv)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    chkValor_serv.Value = 1
    txtValorCentro_Serv.SetFocus
    Exit Sub
End If

'Verifica se o valor do centro de custo já ultrapassou o valor do item
Qtde = 0
Qtd = IIf(txtValorCentro_Serv = "", 0, txtValorCentro_Serv)
qt = txtValorTotal_serv
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(valor) as qtde from Compras_pedido_lista_custo where IDLista = " & txtIDLista_serv & " and ID <> " & txtIDCentro_serv, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
End If
TBAbrir.Close
If Format((Qtde + Qtd), "###,##0.00") > qt Then
    USMsgBox ("Não é permitido salvar, pois o valor do centro de custo ultrapassou o valor total do serviço."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
ID_CC = 0
'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If Novo_PC2_Custo = False Then ID_CC = Lista_custoServ.SelectedItem.ListSubItems(5)

    If ID_CC <> Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex) Then
        Formulario = "Compras/Autorização de centro de custo sem previsão"
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select ID_PC from projproduto where desenho = '" & txtCodigo & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            If USMsgBox("O produto não possui conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
        Else
            Set TBCQ = CreateObject("adodb.recordset")
            TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
            If TBCQ.EOF = True Then
                If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
            Else
                Permitido = True
            End If
            TBCQ.Close
        End If
        TBProduto.Close
        If Permitido = False Then Exit Sub
    End If
End If
TBTempo.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_pedido_lista where IDlista = " & txtIDLista_serv, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Novo_PC2_Custo = True Then TextoFiltro = "ID_CC = " & Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex) Else TextoFiltro = "ID_CC = " & Lista_custoServ.SelectedItem.ListSubItems(5)
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from Compras_pedido_lista_custo where IDlista = " & TBAbrir!IDlista & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar!IDpedido = txtIDPedido
        TBGravar!IDlista = TBAbrir!IDlista
        TBGravar!Responsavel = pubUsuario
        TBGravar!Data = Date
        Evento = "Novo centro de custo do serviço"
    Else
        If txtResponsavel_aprovacao <> "" Then
            If txtResponsavel_aprovacao <> pubUsuario Then
                If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "alterar", True, True) = False Then Exit Sub
            End If
        Else
            If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "alterar", True, True) = False Then Exit Sub
        End If
        If FunVerifSatus("alterar este centro de custo", True) = False Then Exit Sub
        If FunVerifSatusProdServ(txtStatus_serv, "alterar este centro de custo", True, False) = False Then Exit Sub
        
        Evento = "Alterar centro de custo do serviço"
    End If
    TBGravar!ID_CC = Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex)
    TBGravar!Percentual = txtPercentualCentro_Serv
    
    Qtde = txtPercentualCentro_Serv
    Qtd = TBAbrir!preco_total
    qt = (Qtd * Qtde) / 100
    TBGravar!valor = Format(qt, "###,##0.00")
    
    TBGravar.Update
    TBGravar.Close
    
    '==================================
    Modulo = "Compras/Pedido"
    ID_documento = TBAbrir!IDlista
    Documento = "Nº pedido: " & txtPedido
    Documento1 = "Cód. interno: " & TBAbrir!Desenho & " - Centro de custo: " & Cmb_centro_servico
    ProcGravaEvento
    '==================================
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID from Compras_pedido_lista_custo where IDlista = " & txtIDLista_serv & " and ID_CC = " & Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtIDCentro = TBAbrir!ID
    End If
    TBAbrir.Close

    ProcCarregaLista_CustoServ
    If Novo_PC2_Custo = True Then
        USMsgBox ("Novo centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        If Lista_custoServ.ListItems.Count <> 0 And CodigoLista4 <> 0 Then
            Lista_custoServ.SelectedItem = Lista_custoServ.ListItems(CodigoLista4)
            Lista_custoServ.SetFocus
        End If
    End If
    Novo_PC2_Custo = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirServ_custo()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_custoServ
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) centro(s) de custo do serviço " & txtCodigo & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select * from Compras_pedido_lista_custo where ID = " & .ListItems(InitFor) & " and ID_requisicao <> 0 and ID_requisicao is not null", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then
                TBCompras_Lista!IDpedido = 0
                TBCompras_Lista!valor = 0
                TBCompras_Lista!Percentual = 0
                TBCompras_Lista.Update
            Else
                Conexao.Execute "DELETE from Compras_pedido_lista_custo WHERE ID = " & .ListItems(InitFor)
            End If
            TBCompras_Lista.Close
            
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Excluir centro de custo do serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº pedido: " & txtPedido
            If .ListItems(InitFor).SubItems(1) <> "" Then CC = .ListItems(InitFor).SubItems(1) & " - " & .ListItems(InitFor).SubItems(2) Else CC = .ListItems(InitFor).SubItems(2)
            Documento1 = "Cód. interno: " & txtCodigo & " - Centro de custo: " & CC
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centro(s) de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Centro(s) de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCustoServ
    Frame1(8).Enabled = False
    ProcCarregaLista_CustoServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposCustoServ()
On Error GoTo tratar_erro

txtIDCentro_serv = 0
ProcCarregaComboSetor Cmb_centro_servico, "Setor is not null and DtBloq IS NULL and (Consolidacao = 'False' or Consolidacao is null)", txtCodigo, False, True, False, "", True, False
txtValorCentro_Serv = ""
chkValor_serv.Value = 0
chkPercentual_serv.Value = 0
txtValorCentro_Serv.Locked = True
txtValorCentro_Serv.TabStop = False
With txtPercentualCentro_Serv
    .Text = ""
    .Locked = True
    .TabStop = False
End With
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1:
        If SSTab4.Tab = 0 Then ProcFiltrar_Necessidade Else ProcFiltrar_Solicitacao
    Case 2: ProcGerarPed
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcCopiar
    
    If txtIDPedido <> 0 Then
    ProcSalvarPedidoWEB (Int(txtIDPedido))
    End If
    
    Case 9: ProcValidarRegistros listapedido, "Compras/Pedido"
    Case 10: ProcValidarRegistros listapedido, "Compras/Pedido/Aprovar"
    Case 11: ProcEnviarEmail
    Case 12: ProcExportarExcel
    Case 13: procAtualiza
    Case 15: ProcAjuda
    Case 16: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvarComercial
    Case 2: ProcImprimir
    Case 3: ProcAnterior
    Case 4: ProcProximo
    Case 5: ProcFinanceiro
    Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoTab_Produto
    Case 2: ProcSalvarTab_Produto
    Case 3: ProcExcluirTab_Produto
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: procCalculadora
    Case 8: ProcAlterarStatusItem
    Case 9: ProcAlteracoes
    Case 10: ProcCopiar_CC
    Case 11: ProcNecess_solici
    Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAprovacaoItem()
On Error GoTo tratar_erro

If Listprod.ListItems.Count > 0 Then

If Listprod.SelectedItem.ListSubItems.Item(13).Text = "COMPRADO" Then
    If USMsgBox("Deseja realmente cancelar a aprovação desse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        StrSql = "Update compras_Pedido_Lista set Status_item = 'AGUARDANDO APROVAÇÃO' WHERE idLista = '" & Listprod.SelectedItem & "'"
        Conexao.Execute (StrSql)
        ProcAtualizalista
        Exit Sub
    End If
End If

If Listprod.SelectedItem.ListSubItems.Item(13).Text = "AGUARDANDO APROVAÇÃO" Then
    If USMsgBox("Deseja realmente aprovar a compra desse item?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        StrSql = "Update compras_Pedido_Lista set Status_item = 'N_RECEBIDO' WHERE idLista = '" & Listprod.SelectedItem & "'"
        Conexao.Execute (StrSql)
        ProcAtualizalista
        Exit Sub
    End If
End If

End If
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoTab_Servico
    Case 2: ProcSalvarTab_Servico
    Case 3: ProcExcluirTab_Servico
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcAlterarStatusServico
    Case 8: ProcAlteracoes
    Case 9: ProcCopiar_CC_Serv
    Case 10: ProcNecess_solici
    Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar6_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoEscopo
    Case 2: ProcLocalizarEscopo
    Case 3: ProcSalvarEscopo
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlteracoes()
On Error GoTo tratar_erro

Permitido = True
If SSTab1.Tab = 3 Then
    TextoPadrao = "produtos"
    Sit_REG = 1
    If txtcodproduto = 0 Then Permitido = False
Else
    TextoPadrao = "serviço"
    Sit_REG = 2
    If txtcodproduto_serv = 0 Then Permitido = False
End If
If Permitido = False Then
    USMsgBox ("Informe o " & TextoPadrao & " antes de cadastrar as alterações."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Compras_Pedido = True
Vendas_Proposta = False
Vendas_PI = False
frmVendas_PI_alteracoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNecess_solici()
On Error GoTo tratar_erro

Permitido = True
If SSTab1.Tab = 3 Then
    TextoPadrao = "produtos"
    'Sit_REG = 1
    If txtcodproduto = 0 Then Permitido = False
Else
    TextoPadrao = "serviço"
    'Sit_REG = 2
    If txtcodproduto_serv = 0 Then Permitido = False
End If
If Permitido = False Then
    USMsgBox ("Informe o " & TextoPadrao & " antes de cadastrar as alterações."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmCompras_Pedido_NecessSolici.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcFinanceiro()
On Error GoTo tratar_erro

If txtStatus = "AGUARDANDO APROVAÇÃO" Then
    USMsgBox ("Não é permitido gerar o financeiro deste pedido, pois o mesmo ainda não foi aprovado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "enviar para o financeiro"
If txtTotalPedido = "" Or txtTotalPedido = "0,00" Then
    NomeCampo = "o valor do pedido"
    ProcVerificaAcao
    Exit Sub
End If
frmCompras_pedido_MenuFinanceiro.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With txtEspecificacoes
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercial
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
cmbReferencia.Visible = True
txtreferencia.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCampos_serv()
On Error GoTo tratar_erro

With txtDescricao_serv
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercialServ
    .Locked = True
    .TabStop = False
End With
With cmbUn_serv
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com_serv
    .Locked = True
    .TabStop = False
End With
With cmbFamilia_serv
    .Locked = True
    .TabStop = False
End With
cmbreferencia_serv.Visible = True
txtReferencia_serv.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procliberacampos()
On Error GoTo tratar_erro

With txtEspecificacoes
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercial
    .Locked = False
    .TabStop = True
End With
With cmbun
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
If chkAuto.Value = 1 Or chkManual.Value = 1 Then
    cmbReferencia.Visible = False
    txtreferencia.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLiberaCampos_serv()
On Error GoTo tratar_erro

With txtDescricao_serv
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercialServ
    .Locked = False
    .TabStop = True
End With
With cmbUn_serv
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com_serv
    .Locked = False
    .TabStop = True
End With
With cmbFamilia_serv
    .Locked = False
    .TabStop = True
End With
If chkAuto_serv.Value = 1 Or chkManual_serv.Value = 1 Then
    cmbreferencia_serv.Visible = False
    txtReferencia_serv.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosCFOPProdServ(ID_CFOP As Long, Prod As Boolean)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Prod = True Then
        Txt_ID_CFOP_prod = TBAbrir!IDCountCfop
        txtCFOP_prod = IIf(IsNull(TBAbrir!ID_CFOP), "", TBAbrir!ID_CFOP)
        Txt_natureza_operacao_prod = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
    Else
        Txt_ID_CFOP_serv = TBAbrir!IDCountCfop
        txtCFOP_serv = IIf(IsNull(TBAbrir!ID_CFOP), "", TBAbrir!ID_CFOP)
        txtNatureza_operacao_serv = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaCST()
On Error GoTo tratar_erro

If Txt_ID_CFOP_prod = "" Then Exit Sub
Cmb_CST_ICMS.Clear
Cmb_CST_ICMS.AddItem ""
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select * from tbl_NaturezaOperacao_CST where ID_CFOP = " & Txt_ID_CFOP_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    'CST de ICMS
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CST_ICMS from tbl_NaturezaOperacao_CST where ID_CFOP = " & Txt_ID_CFOP_prod & " group by CST_ICMS", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            If IsNull(TBCFOP!CST_ICMS) = False And TBCFOP!CST_ICMS <> "" Then Cmb_CST_ICMS.AddItem TBCFOP!CST_ICMS
            TBCFOP.MoveNext
        Loop
    End If
Else
    With Cmb_CST_ICMS
        .AddItem "000"
        .AddItem "010"
        .AddItem "0101"
        .AddItem "0102"
        .AddItem "0103"
        .AddItem "020"
        .AddItem "0201"
        .AddItem "0202"
        .AddItem "0203"
        .AddItem "0300"
        .AddItem "040"
        .AddItem "0400"
        .AddItem "041"
        .AddItem "050"
        .AddItem "0500"
        .AddItem "051"
        .AddItem "060"
        .AddItem "070"
        .AddItem "090"
        .AddItem "0900"
        
        .AddItem "100"
        .AddItem "110"
        .AddItem "1101"
        .AddItem "1102"
        .AddItem "1103"
        .AddItem "120"
        .AddItem "1201"
        .AddItem "1202"
        .AddItem "1203"
        .AddItem "1300"
        .AddItem "140"
        .AddItem "1400"
        .AddItem "141"
        .AddItem "150"
        .AddItem "1500"
        .AddItem "151"
        .AddItem "160"
        .AddItem "170"
        .AddItem "190"
        .AddItem "1900"
        
        .AddItem "200"
        .AddItem "210"
        .AddItem "2101"
        .AddItem "2102"
        .AddItem "2103"
        .AddItem "220"
        .AddItem "2201"
        .AddItem "2202"
        .AddItem "2203"
        .AddItem "2300"
        .AddItem "240"
        .AddItem "2400"
        .AddItem "241"
        .AddItem "250"
        .AddItem "2500"
        .AddItem "251"
        .AddItem "260"
        .AddItem "270"
        .AddItem "290"
        .AddItem "2900"
        
        .AddItem "300"
        .AddItem "310"
        .AddItem "3101"
        .AddItem "3102"
        .AddItem "3103"
        .AddItem "320"
        .AddItem "3201"
        .AddItem "3202"
        .AddItem "3203"
        .AddItem "3300"
        .AddItem "340"
        .AddItem "3400"
        .AddItem "341"
        .AddItem "350"
        .AddItem "3500"
        .AddItem "351"
        .AddItem "360"
        .AddItem "370"
        .AddItem "390"
        .AddItem "3900"
        
        .AddItem "400"
        .AddItem "410"
        .AddItem "4101"
        .AddItem "4102"
        .AddItem "4103"
        .AddItem "420"
        .AddItem "4201"
        .AddItem "4202"
        .AddItem "4203"
        .AddItem "4300"
        .AddItem "440"
        .AddItem "4400"
        .AddItem "441"
        .AddItem "450"
        .AddItem "4500"
        .AddItem "451"
        .AddItem "460"
        .AddItem "470"
        .AddItem "490"
        .AddItem "4900"
        
        .AddItem "500"
        .AddItem "510"
        .AddItem "5101"
        .AddItem "5102"
        .AddItem "5103"
        .AddItem "520"
        .AddItem "5201"
        .AddItem "5202"
        .AddItem "5203"
        .AddItem "5300"
        .AddItem "540"
        .AddItem "5400"
        .AddItem "541"
        .AddItem "550"
        .AddItem "5500"
        .AddItem "551"
        .AddItem "560"
        .AddItem "570"
        .AddItem "590"
        .AddItem "5900"
        
        .AddItem "600"
        .AddItem "610"
        .AddItem "6101"
        .AddItem "6102"
        .AddItem "6103"
        .AddItem "620"
        .AddItem "6201"
        .AddItem "6202"
        .AddItem "6203"
        .AddItem "6300"
        .AddItem "640"
        .AddItem "6400"
        .AddItem "641"
        .AddItem "650"
        .AddItem "6500"
        .AddItem "651"
        .AddItem "660"
        .AddItem "670"
        .AddItem "690"
        .AddItem "6900"
        
        .AddItem "700"
        .AddItem "710"
        .AddItem "7101"
        .AddItem "7102"
        .AddItem "7103"
        .AddItem "720"
        .AddItem "7201"
        .AddItem "7202"
        .AddItem "7203"
        .AddItem "7300"
        .AddItem "740"
        .AddItem "7400"
        .AddItem "741"
        .AddItem "750"
        .AddItem "7500"
        .AddItem "751"
        .AddItem "760"
        .AddItem "770"
        .AddItem "790"
        .AddItem "7900"
    End With
End If
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCorrigeStatusPedido(StatusPed As String)
On Error GoTo tratar_erro

FunCorrigeStatusPedido = StatusPed
Select Case StatusPed
    Case "ABERTO": FunCorrigeStatusPedido = "COMPRADO"
    Case "PARCIAL": FunCorrigeStatusPedido = "RECEBIDO PARCIAL"
    Case "ENCERRADO": FunCorrigeStatusPedido = "RECEBIDO"
End Select

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCopiar_CC()
On Error GoTo tratar_erro

If Novo_Novo_PC1_Custo = True Then
    USMsgBox ("Informe o centro de custo na lista antes de copiar."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "copiar", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "copiar", True, True) = False Then Exit Sub
End If
If FunVerifSatus("copiar este centro de custo", True) = False Then Exit Sub
If FunVerifSatusProdServ(txtstatus_item, "copiar este centro de custo", True, True) = False Then Exit Sub

Acao = "copiar"
If Cmb_centro = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Exit Sub
End If
If chkPercentual.Value = 1 Then
    Valor1 = IIf(txtPercentualCentro = "", 0, txtPercentualCentro)
    If Valor1 = 0 Then
        NomeCampo = "o percentual"
        ProcVerificaAcao
        txtPercentualCentro.SetFocus
        Exit Sub
    End If
End If
Valor1 = IIf(txtValorCentro = "", 0, txtValorCentro)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    chkValor.Value = 1
    txtValorCentro.SetFocus
    Exit Sub
End If

'Verifica se o valor do centro de custo já ultrapassou o valor do item
Qtde = 0
Qtd = IIf(txtValorCentro = "", 0, txtValorCentro)
qt = txtvlrTotal
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(valor) as qtde from Compras_pedido_lista_custo where IDLista = " & TXTIDLista & " and id_CC <> " & txtIDCentro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
End If
TBAbrir.Close
If Format((Qtde + Qtd), "###,##0.00") > qt Then
    USMsgBox ("Não é permitido salvar, pois o valor do centro de custo ultrapassou o valor total do produto."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido1 = False
If USMsgBox("Deseja copiar o centro de custo para todos produtos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_pedido_lista where IDPedido = " & txtIDPedido & " and (Status_Item = 'AGUARDANDO APROVAÇÃO' or Status_Item = 'N_RECEBIDO') and Tipo = 'P' order by idlista", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
            If Permitido1 = False Then
                Permitido = False
                Set TBTempo = CreateObject("adodb.recordset")
                TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBTempo.EOF = False Then
                    Formulario = "Compras/Autorização de centro de custo sem previsão"
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select ID_PC from projproduto where desenho = '" & IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho) & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = True Then
                        If USMsgBox("Existe(m) produto(s) sem conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                    Else
                        Set TBCQ = CreateObject("adodb.recordset")
                        TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & Cmb_centro.ItemData(Cmb_centro.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCQ.EOF = True Then
                            If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                        Else
                            Permitido = True
                        End If
                        TBCQ.Close
                    End If
                    TBProduto.Close
                    If Permitido = False Then Exit Sub
                End If
                TBTempo.Close
            End If
        
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "select * from Compras_pedido_lista_custo where IDlista = " & TBAbrir!IDlista & " and ID_CC = " & Lista_custo.SelectedItem.ListSubItems(5), Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                TBGravar!IDpedido = txtIDPedido
                TBGravar!IDlista = TBAbrir!IDlista
                TBGravar!Responsavel = pubUsuario
                TBGravar!Data = Date
                Evento = "Novo centro de custo do produto"
            Else
                Evento = "Alterar centro de custo do produto"
            End If
            TBGravar!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex)
            TBGravar!Percentual = txtPercentualCentro
            
            Qtde = txtPercentualCentro
            Qtd = TBAbrir!preco_total
            qt = (Qtd * Qtde) / 100
            TBGravar!valor = Format(qt, "###,##0.00")
            
            TBGravar.Update
            TBGravar.Close
            
            '==================================
            Modulo = "Compras/Pedido"
            ID_documento = TBAbrir!IDlista
            Documento = "Nº pedido: " & txtPedido
            Documento1 = "Cód. interno: " & TBAbrir!Desenho & " - Centro de custo: " & Cmb_centro
            ProcGravaEvento
            '==================================
        
            TBAbrir.MoveNext
        Loop
        
        ProcCarregaLista_Custo
        USMsgBox ("Centro de custo copiado com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiar_CC_Serv()
On Error GoTo tratar_erro

If Novo_PC2_Custo = True Then
    USMsgBox ("Informe o centro de custo na lista antes de copiar."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If txtResponsavel_aprovacao <> "" Then
    If txtResponsavel_aprovacao <> pubUsuario Then
        If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "copiar", True, True) = False Then Exit Sub
    End If
Else
    If FunVerificaRegistroValidado("Compras_pedido", "IDPedido = " & txtIDPedido, "pedido de compra", "este centro de custo", "copiar", True, True) = False Then Exit Sub
End If
If FunVerifSatus("copiar este centro de custo", True) = False Then Exit Sub
If FunVerifSatusProdServ(txtStatus_serv, "copiar este centro de custo", True, False) = False Then Exit Sub

Acao = "copiar"
If Cmb_centro_servico = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Exit Sub
End If
If chkPercentual_serv.Value = 1 Then
    Valor1 = IIf(txtPercentualCentro_Serv = "", 0, txtPercentualCentro_Serv)
    If Valor1 = 0 Then
        NomeCampo = "o percentual"
        ProcVerificaAcao
        txtPercentualCentro_Serv.SetFocus
        Exit Sub
    End If
End If
Valor1 = IIf(txtValorCentro_Serv = "", 0, txtValorCentro_Serv)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    chkValor.Value = 1
    txtValorCentro_Serv.SetFocus
    Exit Sub
End If

'Verifica se o valor do centro de custo já ultrapassou o valor do item
Qtde = 0
Qtd = IIf(txtValorCentro_Serv = "", 0, txtValorCentro_Serv)
qt = txtValorTotal_serv
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(valor) as qtde from Compras_pedido_lista_custo where IDLista = " & txtIDLista_serv & " and id <> " & txtIDCentro_serv, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
End If
TBAbrir.Close
If Format((Qtde + Qtd), "###,##0.00") > qt Then
    USMsgBox ("Não é permitido salvar, pois o valor do centro de custo ultrapassou o valor total do serviço."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido1 = False
If USMsgBox("Deseja copiar o centro de custo para todos produtos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_pedido_lista where IDPedido = " & txtIDPedido & " and (Status_Item = 'AGUARDANDO APROVAÇÃO' or Status_Item = 'N_RECEBIDO') and Tipo = 'S' order by idlista", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
            If Permitido1 = False Then
                Permitido = False
                Set TBTempo = CreateObject("adodb.recordset")
                TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBTempo.EOF = False Then
                    Formulario = "Compras/Autorização de centro de custo sem previsão"
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select ID_PC from projproduto where desenho = '" & IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho) & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = True Then
                        If USMsgBox("Existe(m) serviço(s) sem conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                    Else
                        Set TBCQ = CreateObject("adodb.recordset")
                        TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCQ.EOF = True Then
                            If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                        Else
                            Permitido = True
                        End If
                        TBCQ.Close
                    End If
                    TBProduto.Close
                    If Permitido = False Then Exit Sub
                End If
                TBTempo.Close
            End If
        
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "select * from Compras_pedido_lista_custo where IDlista = " & TBAbrir!IDlista & " and ID_CC = " & Lista_custoServ.SelectedItem.ListSubItems(5), Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                TBGravar!IDpedido = txtIDPedido
                TBGravar!IDlista = TBAbrir!IDlista
                TBGravar!Responsavel = pubUsuario
                TBGravar!Data = Date
                Evento = "Novo centro de custo do serviço"
            Else
                Evento = "Alterar centro de custo do serviço"
            End If
            TBGravar!ID_CC = Cmb_centro_servico.ItemData(Cmb_centro_servico.ListIndex)
            TBGravar!Percentual = txtPercentualCentro_Serv
            
            Qtde = txtPercentualCentro_Serv
            Qtd = TBAbrir!preco_total
            qt = (Qtd * Qtde) / 100
            TBGravar!valor = Format(qt, "###,##0.00")
            
            TBGravar.Update
            TBGravar.Close
            
            '==================================
            Modulo = "Compras/Pedido"
            ID_documento = TBAbrir!IDlista
            Documento = "Nº pedido: " & txtPedido
            Documento1 = "Cód. interno: " & TBAbrir!Desenho & " - Centro de custo: " & Cmb_centro_servico
            ProcGravaEvento
            '==================================
        
            TBAbrir.MoveNext
        Loop
        
        ProcCarregaLista_CustoServ
        USMsgBox ("Centro de custo copiado com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifSatus(Acao As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifSatus = True
If (txtStatus = "COMPRADO") And txtResponsavel_aprovacao <> pubUsuario Or txtStatus = "RECEBIDO" Or txtStatus = "CANCELADO" Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & ", pois o " & IIf(Right(Acao, 6) = "pedido", "o mesmo", "pedido de compra") & " está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    FunVerifSatus = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function FunVerifSatusProdServ(status As String, Acao As String, MostrarMsg As Boolean, Prod As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifSatusProdServ = True
If (status = "COMPRADO" Or status = "RECEBIDO PARCIAL") And txtResponsavel_aprovacao <> pubUsuario Or status = "RECEBIDO" Or status = "CANCELADO" Then
    If MostrarMsg = True Then
        If Right(Acao, 7) = "produto" Or Right(Acao, 7) = "serviço" Then MsgTexto = "mesmo" Else MsgTexto = ""
        USMsgBox ("Não é permitido " & Acao & ", pois o " & IIf(MsgTexto = "", IIf(Prod = True, "produto", "serviço"), MsgTexto) & " está " & status & "."), vbExclamation, "CAPRIND v5.0"
    End If
    FunVerifSatusProdServ = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcFiltrar_Necessidade()
On Error GoTo tratar_erro

If Opt_PCP.Value = True Then NomeTabela = "Estoque_necessidade_resumido" Else NomeTabela = "Estoque_necessidade_resumido_PIEST"
CamposFiltro = "EN.Codproduto, EN.Desenho, EN.Descricao, EN.Unidade, EN.Unidade_com, EN.Necessidade, EN.Necessidade_estoque"
INNERJOINTEXTO = "Select " & CamposFiltro & " from " & NomeTabela & " EN "
If Cmb_filtrar = "Com necessidade" Then TextoFiltroEstoque = " and EN.Necessidade > 0" Else TextoFiltroEstoque = " and EN.Necessidade_estoque > 0"
TextoFiltroPadrao = "EN.Compras = 'True' and EN.ID_empresa = " & Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex) & TextoFiltroEstoque & " group by " & CamposFiltro & " order by EN.desenho"

If txtTexto_necess.Visible = True And txtTexto_necess <> "" Or cmbTexto_necess.Visible = True And cmbTexto_necess <> "" Then
    If cmbfiltrarpor_necess = "Família" Then
        StrSql_Pedido_Necessidade = INNERJOINTEXTO & " where EN.classe = '" & cmbTexto_necess & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor_necess
            Case "Código interno": TextoFiltro = "EN.Desenho"
            Case "Código de referência": TextoFiltro = "IA.n_referencia"
            Case "Descrição": TextoFiltro = "EN.Descricao"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
        End Select
        StrSql_Pedido_Necessidade = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_necess, Optmeio_necess, Optfim_necess, optIgual_necess, txtTexto_necess) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Pedido_Necessidade = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

'Debug.print StrSql_Pedido_Necessidade

ProcCarregalista_Necessidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Necessidade()
On Error GoTo tratar_erro

If StrSql_Pedido_Necessidade = "" Then Exit Sub
lblRegistros(1).Caption = "Nº de registros: 0"
lblPaginas(1).Caption = "Página: 0 de: 0"
ListaNecessidade.ListItems.Clear
Set TBLISTA_Pedido_Necessidade = CreateObject("adodb.recordset")
TBLISTA_Pedido_Necessidade.Open StrSql_Pedido_Necessidade, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Pedido_Necessidade.EOF = False Then ProcExibePagina_Necessidade (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Necessidade(Pagina)
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear
TBLISTA_Pedido_Necessidade.PageSize = IIf(txtNreg(1) = "", 30, txtNreg(1))
TBLISTA_Pedido_Necessidade.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Pedido_Necessidade.PageSize
ContadorReg = 1

PBLista(0).Min = 0
PBLista(0).Max = FunVerifMaxPBListaPaginacao(TBLISTA_Pedido_Necessidade.RecordCount - IIf(Pagina > 1, (TBLISTA_Pedido_Necessidade.PageSize * (Pagina - 1)), 0), TBLISTA_Pedido_Necessidade.PageSize)
PBLista(0).Value = 1
Contador = 0
Do While TBLISTA_Pedido_Necessidade.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaNecessidade.ListItems.Add(, , TBLISTA_Pedido_Necessidade!Codproduto)
        .SubItems(1) = IIf(IsNull(TBLISTA_Pedido_Necessidade!Desenho), "", TBLISTA_Pedido_Necessidade!Desenho)
        .SubItems(2) = IIf(IsNull(TBLISTA_Pedido_Necessidade!Descricao), "", TBLISTA_Pedido_Necessidade!Descricao)
        .SubItems(3) = IIf(IsNull(TBLISTA_Pedido_Necessidade!Unidade_com), "", TBLISTA_Pedido_Necessidade!Unidade_com)
        If TBLISTA_Pedido_Necessidade!Unidade <> TBLISTA_Pedido_Necessidade!Unidade_com Then
            If Cmb_filtrar = "Com necessidade" Then qt = Format(TBLISTA_Pedido_Necessidade!Necessidade, "###,##0.0000") Else qt = Format(TBLISTA_Pedido_Necessidade!Necessidade_estoque, "###,##0.0000")
            If FunVerifUNConversao(TBLISTA_Pedido_Necessidade!Unidade, TBLISTA_Pedido_Necessidade!Unidade_com) = True Then
                Qtde = FunConverteUN(TBLISTA_Pedido_Necessidade!Unidade_com, TBLISTA_Pedido_Necessidade!Unidade, qt, TBLISTA_Pedido_Necessidade!Desenho)
                .SubItems(4) = Format(Qtde, "###,##0.0000")
            Else
                If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLISTA_Pedido_Necessidade!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLISTA_Pedido_Necessidade!Necessidade_estoque, "###,##0.0000")
            End If
        Else
            If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLISTA_Pedido_Necessidade!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLISTA_Pedido_Necessidade!Necessidade_estoque, "###,##0.0000")
        End If
       ' .SubItems(5) = IIf(IsNull(TBLISTA_Pedido_Necessidade!Part_number), "", TBLISTA_Pedido_Necessidade!Part_number)
       ' .SubItems(6) = IIf(IsNull(TBLISTA_Pedido_Necessidade!Fabricante), "", TBLISTA_Pedido_Necessidade!Fabricante)
        If Cmb_filtrar = "Com necess. estoque" Then NReal = Format(TBLISTA_Pedido_Necessidade!Necessidade_estoque, "###,##0.0000") Else NReal = Format(TBLISTA_Pedido_Necessidade!Necessidade, "###,##0.0000")
        If NReal > 0 Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .ListSubItems(3).ForeColor = vbRed
            .ListSubItems(4).ForeColor = vbRed
            '.ListSubItems(5).ForeColor = vbRed
        End If
    End With
    TBLISTA_Pedido_Necessidade.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista(0).Value = Contador
Loop
lblRegistros(1).Caption = "Nº de registros: " & TBLISTA_Pedido_Necessidade.RecordCount
If TBLISTA_Pedido_Necessidade.AbsolutePage = adPosBOF Then
   lblPaginas(1).Caption = "Página: 1 de: " & TBLISTA_Pedido_Necessidade.PageCount
ElseIf TBLISTA_Pedido_Necessidade.AbsolutePage = adPosEOF Then
        lblPaginas(1).Caption = "Página: " & TBLISTA_Pedido_Necessidade.PageCount & " de: " & TBLISTA_Pedido_Necessidade.PageCount
    Else
        lblPaginas(1).Caption = "Página: " & TBLISTA_Pedido_Necessidade.AbsolutePage - 1 & " de: " & TBLISTA_Pedido_Necessidade.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Solicitacao()
On Error GoTo tratar_erro

CamposFiltro = "CR.ID_requisicao, CR.Requisicaotexto, CPL.IdLista, CPL.Status_Item, CPL.desenho, CPL.descricao, CPL.Un, CPL.Unidade_com, CPL.quant_req, CPL.detalheitem, CPL.prazoreq, CPL.obs"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (Compras_requisicao CR INNER JOIN Compras_pedido_lista CPL ON CR.ID_Requisicao = CPL.ID_Requisicao) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CPL.codproduto"
TextoFiltroPadrao = "CPL.status_item = 'REQUISIT.' and CR.Status = 'LIBERADA' group by " & CamposFiltro & " order by CR.ID_requisicao"

If txtTexto_sol.Visible = True And txtTexto_sol <> "" Or cmbTexto_sol.Visible = True And cmbTexto_sol <> "" Or Txtprazo_sol.Visible = True Then
    If cmbfiltrarpor_sol = "Família" Then
        StrSql_Pedido_Solicitacao = INNERJOINTEXTO & " where CPL.Familia = '" & cmbTexto_sol & "' and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
            StrSql_Pedido_Solicitacao = INNERJOINTEXTO & " where CPL.Prazoreq = '" & Format(Txtprazo_sol.Value, "Short Date") & "' and " & TextoFiltroPadrao
        Else
            Select Case cmbfiltrarpor_sol
                Case "Solicitação": TextoFiltro = "CR.Requisicaotexto"
                Case "Código interno": TextoFiltro = "CPL.desenho"
                Case "Descrição": TextoFiltro = "CPL.descricao"
                Case "Descrição comercial": TextoFiltro = "CPL.descricao_comercial"
                Case "Detalhe": TextoFiltro = "CPL.Detalheitem"
                Case "Part number": TextoFiltro = "PFAB.Part_number"
            End Select
            StrSql_Pedido_Solicitacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_sol, optMeio_sol, optFim_sol, optIgual_sol, txtTexto_sol) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Pedido_Solicitacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Solicitacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Solicitacao()
On Error GoTo tratar_erro

lblRegistros(2).Caption = "Nº de reg.: 0"
lblPaginas(2).Caption = "Página: 0 de: 0"
Lista_solicitados.ListItems.Clear
If StrSql_Pedido_Solicitacao = "" Then Exit Sub
Set TBLISTA_Pedido_Solicitacao = CreateObject("adodb.recordset")
TBLISTA_Pedido_Solicitacao.Open StrSql_Pedido_Solicitacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Pedido_Solicitacao.EOF = False Then ProcExibePagina_Solicitacao (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Solicitacao(Pagina)
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear
TBLISTA_Pedido_Solicitacao.PageSize = IIf(txtNreg(2) = "", 30, txtNreg(2))
TBLISTA_Pedido_Solicitacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Pedido_Solicitacao.PageSize
ContadorReg = 1

PBLista(0).Min = 0
PBLista(0).Max = FunVerifMaxPBListaPaginacao(TBLISTA_Pedido_Solicitacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Pedido_Solicitacao.PageSize * (Pagina - 1)), 0), TBLISTA_Pedido_Solicitacao.PageSize)
PBLista(0).Value = 1
Contador = 0
Do While TBLISTA_Pedido_Solicitacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_solicitados.ListItems
        .Add , , TBLISTA_Pedido_Solicitacao!IDlista
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Status_Item), "", TBLISTA_Pedido_Solicitacao!Status_Item)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Requisicaotexto), "", TBLISTA_Pedido_Solicitacao!Requisicaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Desenho), "", TBLISTA_Pedido_Solicitacao!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Descricao), "", TBLISTA_Pedido_Solicitacao!Descricao)
        
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Un), "", TBLISTA_Pedido_Solicitacao!Un)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Unidade_com), "", TBLISTA_Pedido_Solicitacao!Unidade_com)
        If TBLISTA_Pedido_Solicitacao!Un <> TBLISTA_Pedido_Solicitacao!Unidade_com Then valor = FunConversaoFinalUn(TBLISTA_Pedido_Solicitacao!Un, TBLISTA_Pedido_Solicitacao!Unidade_com, TBLISTA_Pedido_Solicitacao!quant_req, TBLISTA_Pedido_Solicitacao!Desenho, True) Else valor = TBLISTA_Pedido_Solicitacao!quant_req
        .Item(.Count).SubItems(7) = FunFormataCasasDecimais(4, valor)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!quant_req), "", FunFormataCasasDecimais(4, TBLISTA_Pedido_Solicitacao!quant_req))
        
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!detalheitem), "", TBLISTA_Pedido_Solicitacao!detalheitem)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!prazoreq), "", Format(TBLISTA_Pedido_Solicitacao!prazoreq, "dd/mm/yy"))
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Pedido_Solicitacao!Obs), "", TBLISTA_Pedido_Solicitacao!Obs)
    End With
    ContadorReg = ContadorReg + 1
    TBLISTA_Pedido_Solicitacao.MoveNext
    Contador = Contador + 1
    PBLista(0).Value = Contador
Loop
lblRegistros(2).Caption = "Nº de reg.: " & TBLISTA_Pedido_Solicitacao.RecordCount
If TBLISTA_Pedido_Solicitacao.AbsolutePage = adPosBOF Then
   lblPaginas(2).Caption = "Pág.: 1 de: " & TBLISTA_Pedido_Solicitacao.PageCount
ElseIf TBLISTA_Pedido_Solicitacao.AbsolutePage = adPosEOF Then
        lblPaginas(2).Caption = "Pág.: " & TBLISTA_Pedido_Solicitacao.PageCount & " de: " & TBLISTA_Pedido_Solicitacao.PageCount
    Else
        lblPaginas(2).Caption = "Pág.: " & TBLISTA_Pedido_Solicitacao.AbsolutePage - 1 & " de: " & TBLISTA_Pedido_Solicitacao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarPed()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With ListaNecessidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente gerar pedido deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                ProcLimpar
                ProcLimparTudo
                ProcConfVariaveisLocForn False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
                Novo_PC = True
                Sit_REG = 1
                Permitido2 = True
                FrmCompras_localizafornecedor.Show 1
                If Permitido2 = False Then
                    Novo_PC = False
                    Exit Sub
                End If
                ProcNovoPedido
                
                If USMsgBox("Algum produto/serviço selecionado será adicionado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
            End If
            Permitido = True
            
            IDlista = .ListItems.Item(InitFor)
            Desenho = .ListItems(InitFor).SubItems(1)
            If Permitido1 = True Then
                Compras_Pedido = True
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = False
                Qtde = .ListItems(InitFor).SubItems(4)
                Sit_Data = 1
                Permitido2 = True
                Sit_REG = 1
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then
                    valor = .ListItems(InitFor).SubItems(4)
                    ProcNovo_Necessidade Opt_vendas
                End If
            Else
                valor = .ListItems(InitFor).SubItems(4)
                ProcNovo_Necessidade Opt_vendas
            End If
        End If
    Next InitFor
End With

With Lista_solicitados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente gerar pedido deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                ProcLimpar
                ProcLimparTudo
                
                ProcConfVariaveisLocForn False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
                Novo_PC = True
                Sit_REG = 1
                Permitido2 = True
                FrmCompras_localizafornecedor.Show 1
                If Permitido2 = False Then
                    Novo_PC = False
                    Exit Sub
                End If
                ProcNovoPedido
                
                If Permitido1 = False Then
                    If USMsgBox("Algum produto/serviço selecionado será adicionado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
                End If
            End If
            Permitido = True
            
            IDlista = .ListItems.Item(InitFor)
            Desenho = .ListItems(InitFor).SubItems(3)
            If Permitido1 = True Then
                Compras_Pedido = True
                Vendas_PI = False
                Compras_Cotacao = False
                Faturamento = False
                Qtde = .ListItems(InitFor).SubItems(8)
                Sit_Data = 2
                Permitido2 = True
                Sit_REG = 1
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then
                    valor = .ListItems(InitFor).SubItems(8)
                    ProcAlterar_Solicitacao False
                End If
            Else
                valor = .ListItems(InitFor).SubItems(8)
                ProcAlterar_Solicitacao False
            End If
            
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de gerar o pedido."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    USMsgBox ("Novo pedido cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Frame1(0).Enabled = True
    Frame1(16).Enabled = True
    Novo_PC = False
    
    ProcCarregalista_Necessidade
    ProcCarregalista_Solicitacao
    
    Sql_Pedido_Localizar = "Select CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado, CP.dbl_valor_total from Compras_pedido CP LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao where CP.IDpedido = " & txtIDPedido
    ProcAtualizalistapedido (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoPedido()
On Error GoTo tratar_erro

Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select * from compras_pedido", Conexao, adOpenKeyset, adLockOptimistic
TBCompras_Pedido.AddNew
txtPedido = FunCriarNovoNumero
TBCompras_Pedido!Pedido = txtPedido
TBCompras_Pedido!Responsavel = pubUsuario
TBCompras_Pedido!Data = Date
TBCompras_Pedido!Status_pedido = "AGUARDANDO APROVAÇÃO"
ProcEnviaDados
TBCompras_Pedido.Update
txtIDPedido = TBCompras_Pedido!IDpedido
ProcGravarDCForn
'==================================
Modulo = "Compras/Pedido"
Evento = "Novo"
ID_documento = txtIDPedido
Documento = "Nº pedido: " & txtPedido.Text
Documento1 = ""
ProcGravaEvento
'==================================
ProcPuxaDados
TBCompras_Pedido.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunCriarNovoNumero() As String
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Pedido from compras_pedido where Year(data) = '" & Year(Date) & "' order by IDPedido desc", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Numero = Left(TBTempo!Pedido, Len(TBTempo!Pedido) - 3) + 1
Else
    Numero = 1
End If
TBTempo.Close
Ano = Right(Year(Date), 2)
FunCriarNovoNumero = Numero & "/" & Ano

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcGravarDCForn()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "select * from Compras_comercial where IDpedido = " & txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!IDpedido = txtIDPedido
TBGravar!Moeda = "REAL"
TBGravar!Valor_moeda = 1

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select Tipo_transp, idTransp from Compras_fornecedores where idcliente = " & txtIDfornecedor & " and Transportadora is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    TBGravar!Tipo_transp = IIf(IsNull(TBCompras!Tipo_transp), "", TBCompras!Tipo_transp)
    TBGravar!Idtransporte = IIf(IsNull(TBCompras!idTransp), 0, TBCompras!idTransp)
End If
TBCompras.Close

TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovo_Necessidade(Necess_PI As Boolean)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!IDpedido = txtIDPedido

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select * from projproduto where codproduto = " & IDlista & "", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBGravar!Codproduto = IIf(IsNull(TBProduto!Codproduto), "", TBProduto!Codproduto)
    TBGravar!Desenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
    TBGravar!Descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    TBGravar!Descricao_comercial = IIf(IsNull(TBProduto!descricaotecnica), "", TBProduto!descricaotecnica)
    TBGravar!Un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
    TBGravar!Unidade_com = IIf(IsNull(TBProduto!Unidade_com), "", TBProduto!Unidade_com)
    TBGravar!Familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select PCusto from Projproduto_fornecedor where Codproduto = " & TBProduto!Codproduto & " and idfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        TBGravar!preco_unitario = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto, "###,##0.0000000000"))
    Else
        TBGravar!preco_unitario = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
    End If
    TBFI.Close
    TBGravar!preco_unitario_desconto = TBGravar!preco_unitario
    
    Certificado = ""
    If Necess_PI = False Then NomeTabela = "Estoque_necessidade_detalhado" Else NomeTabela = "Estoque_necessidade_detalhado_PIEST"
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Ped_req, Ordem_req from " & NomeTabela & " where codproduto = " & IDlista & " and Compras = 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by Prazo", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Do While TBFI.EOF = False
            If Certificado = "" Then
                Certificado = IIf(Necess_PI = False, "Ordem: " & TBFI!Ordem_req, "Pedido: " & TBFI!Ped_req)
            Else
                Certificado = Certificado & vbCrLf & IIf(Necess_PI = False, "Ordem: " & TBFI!Ordem_req, "Pedido: " & TBFI!Ped_req)
            End If
            TBFI.MoveNext
        Loop
        TBGravar!NecessSolici = Certificado
    End If
    TBFI.Close
    
    If TBProduto!Tipo = "S" Then
        If IsNull(TBGravar!ID_CFOP) = False Then TBGravar!ID_CFOP = TBProduto!ID_CFOP
        
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select Simples, presumido, Real, Simples1 from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Regime = 0
            If TBFIltro!Simples = True Then Regime = 1
            If TBFIltro!Presumido = True Then Regime = 2
            If TBFIltro!Real = True Then Regime = 3
            If TBFIltro!Simples1 = True Then Regime = 4
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select ISS from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                TBGravar!ISSQN = IIf(IsNull(TBFI!ISS), 0, TBFI!ISS)
            End If
            TBFI.Close
        End If
        
        TBGravar!VlrISSQN = Format(((TBGravar!preco_unitario_desconto * valor) * IIf(IsNull(TBGravar!ISSQN), 0, TBGravar!ISSQN)) / 100, "###,##0.00")
        TBGravar!Tipo = "S"
    Else
        TBGravar!Quant_Comp_PC = FunCalculaQtdePC(TBProduto!Desenho, valor, True, TBGravar!Unidade_com)
        If IsNull(TBProduto!ID_CF) = False Then TBGravar!ID_CF = TBProduto!ID_CF
        
        If IsNull(TBProduto!ID_CFOP) = False Then TBGravar!ID_CFOP = TBProduto!ID_CFOP
        If IsNull(TBProduto!ID_CF) = False Then
            ProcValorImposto txtPedido, IIf(TBProduto!ID_CF = "", 0, TBProduto!ID_CF), IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), 0
            ProcControleImposto IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), IIf(txtIDfornecedor = "", 0, txtIDfornecedor)
            If TemIPI = "SIM" Then TBGravar!IPI = IntIPI Else TBGravar!IPI = 0
            If TemICMS = "SIM" Then TBGravar!ICMS = IntICMS Else TBGravar!ICMS = 0
        End If
        
        TBGravar!VlrIPI = Format(((TBGravar!preco_unitario_desconto * valor) * IIf(IsNull(TBGravar!IPI), 0, TBGravar!IPI)) / 100, "###,##0.00")
        TBGravar!vlrICMS = Format(((TBGravar!preco_unitario_desconto * valor) * IIf(IsNull(TBGravar!ICMS), 0, TBGravar!ICMS)) / 100, "###,##0.00")
        TBGravar!Tipo = "P"
        
        ProcCalculaBCICMSsNecSolic
    End If
End If
TBProduto.Close

TBGravar!Quant_Comp = valor
TBGravar!Status_Item = "AGUARDANDO APROVAÇÃO"
TBGravar!ValorDesconto = 0

TBGravar!preco_total = Format((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) + IIf(IsNull(TBGravar!VlrIPI), 0, TBGravar!VlrIPI), "###,##0.00")
'Calcula quantidade se a unidade for diferente
If TBGravar!Un <> TBGravar!Unidade_com Then
    If FunVerifUNConversao(TBGravar!Un, TBGravar!Unidade_com) = True Then
        TBGravar!Qtde_estoque = FunConverteUN(TBGravar!Un, TBGravar!Unidade_com, TBGravar!Quant_Comp, TBGravar!Desenho)
    Else
        TBGravar!Qtde_estoque = TBGravar!Quant_Comp / FunVerificaTabelaConversaoUnidade(TBGravar!Un, TBGravar!Unidade_com)
    End If
Else
    TBGravar!Qtde_estoque = Null
End If

'TBGravar!Obs_pedido = IIf(IsNull(TBGravar!Obs), "", TBGravar!Obs)
TBGravar.Update

If Necess_PI = True Then
    Valor3 = valor
    'Empenha o pedido de compra para os pedidos de venda mais antigos
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select ENDT.ID, ENDT.Requisitado - ISNULL(CPLE.Qtde_empenho, 0) AS Requisitado from Estoque_necessidade_detalhado ENDT LEFT JOIN Compras_pedido_lista_empenhos CPLE ON CPLE.IDcarteira = ENDT.ID where ENDT.Desenho = '" & Desenho & "' and ENDT.Tipo <> 'OP' and ENDT.Requisitado > ISNULL(CPLE.Qtde_empenho, 0) order by ENDT.Prazo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False And Valor3 > 0
            Set TBCotacao = CreateObject("adodb.recordset")
            TBCotacao.Open "Select * FROM Compras_pedido_lista_empenhos", Conexao, adOpenKeyset, adLockOptimistic
            TBCotacao.AddNew
            TBCotacao!Data = Date
            TBCotacao!Responsavel = pubUsuario
            TBCotacao!IDlista = TBGravar!IDlista
            TBCotacao!IDcarteira = TBCFOP!ID
            If Valor3 >= TBCFOP!Requisitado Then
                TBCotacao!Qtde_empenho = TBCFOP!Requisitado
                Valor3 = Valor3 - TBCFOP!Requisitado
            Else
                TBCotacao!Qtde_empenho = Valor3
                Valor3 = 0
            End If
            TBCotacao.Update
            TBCFOP.MoveNext
        Loop
    End If
End If

'==================================
Modulo = "Compras/Pedido"
Evento = "Novo produto"
ID_documento = TBGravar!IDlista
Documento = "Nº pedido: " & txtPedido
Documento1 = "Cód. interno: " & TBGravar!Desenho
ProcGravaEvento
'==================================
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovo_Solicitacao(Necess_PI As Boolean)
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_pedido_lista where idlista = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA!quant_req = qt
    TBLISTA.Update
    
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!ID_Requisicao = IIf(IsNull(TBLISTA!ID_Requisicao), 0, TBLISTA!ID_Requisicao)
    TBGravar!Codproduto = IIf(IsNull(TBLISTA!Codproduto), "", TBLISTA!Codproduto)
    TBGravar!Tipo = TBLISTA!Tipo
    TBGravar!CODIGO = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
    TBGravar!Status_Item = "AGUARDANDO APROVAÇÃO"
    TBGravar!Un = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
    TBGravar!Unidade_com = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
    TBGravar!Familia = IIf(IsNull(TBLISTA!Familia), "", TBLISTA!Familia)
    TBGravar!solicitado = IIf(IsNull(TBLISTA!solicitado), "", TBLISTA!solicitado)
    TBGravar!setorsolic = IIf(IsNull(TBLISTA!setorsolic), "", TBLISTA!setorsolic)
    TBGravar!Descricao = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
    TBGravar!Descricao_comercial = IIf(IsNull(TBLISTA!Descricao_comercial), "", TBLISTA!Descricao_comercial)
    
    TBGravar!quant_req = valor
    TBGravar!Quant_Comp = valor
    If TBLISTA!Tipo <> "S" Then
        TBGravar!quant_req_PC = FunCalculaQtdePC(TBLISTA!Desenho, valor, True, TBGravar!Unidade_com)
        TBGravar!Quant_Comp_PC = FunCalculaQtdePC(TBLISTA!Desenho, valor, True, TBGravar!Unidade_com)
    End If
    
    TBGravar!Desenho = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
    TBGravar!N_referencia = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
    TBGravar!detalheitem = IIf(IsNull(TBLISTA!detalheitem), "", TBLISTA!detalheitem)
    If TBLISTA!prazoreq <> "" Then
        TBGravar!prazoreq = TBLISTA!prazoreq
        TBGravar!Prazo = TBLISTA!prazoreq
    End If
    TBGravar!Prioridade = TBLISTA!Prioridade
    TBGravar!Remessa = TBLISTA!Remessa
    TBGravar!Obs = TBLISTA!Obs
    TBGravar!Ordem = TBLISTA!Ordem
    TBGravar!OS = TBLISTA!OS
    TBGravar!ID_PC = TBLISTA!ID_PC
    
    TBGravar!IDpedido = txtIDPedido
    TBGravar!ValorDesconto = 0
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select codProduto, Pcusto, ID_CF, ID_CFOP, Tipo from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select PCusto from Projproduto_fornecedor where Codproduto = " & TBProduto!Codproduto & " and idfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TBGravar!preco_unitario = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto, "###,##0.0000000000"))
        Else
            TBGravar!preco_unitario = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
        End If
        TBFI.Close
        If TBLISTA!Remessa = True Then
            valor = Format(FunVerificaQtdeEstoque(TBLISTA!Desenho, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
            If valor > 0 Then TBGravar!preco_unitario = Format(Valor_total / valor, "###,##0.0000000000")
        End If
        TBGravar!preco_unitario_desconto = TBGravar!preco_unitario
        
        If IsNull(TBLISTA!ID_Requisicao) = False Then
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select RequisicaoTexto from Compras_requisicao where ID_requisicao = " & TBLISTA!ID_Requisicao, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then TBGravar!NecessSolici = "Solicitação: " & TBFI!Requisicaotexto & " - Quantidade: " & Format(TBGravar!quant_req, "###,##0.0000") & " - Quantidade PÇ: " & IIf(IsNull(TBGravar!quant_req_PC), 0, TBGravar!quant_req_PC)
            TBFI.Close
        End If
        
        TBGravar!IPI = 0
        TBGravar!ICMS = 0
        If TBProduto!Tipo = "S" Then
            If IsNull(TBGravar!ID_CFOP) = False Then TBGravar!ID_CFOP = TBProduto!ID_CFOP
            
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Regime = 0
                If TBFIltro!Simples = True Then Regime = 1
                If TBFIltro!Presumido = True Then Regime = 2
                If TBFIltro!Real = True Then Regime = 3
                If TBFIltro!Simples1 = True Then Regime = 4
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ISS from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBGravar!ISSQN = IIf(IsNull(TBFI!ISS), 0, TBFI!ISS)
                End If
                TBFI.Close
            End If
            
            TBGravar!VlrISSQN = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!ISSQN), 0, TBGravar!ISSQN)) / 100, "###,##0.00")
        Else
            If IsNull(TBProduto!ID_CF) = False Then TBGravar!ID_CF = TBProduto!ID_CF
            If TBLISTA!Remessa = False Then
                If IsNull(TBProduto!ID_CFOP) = False Then TBGravar!ID_CFOP = TBProduto!ID_CFOP
                If IsNull(TBProduto!ID_CF) = False Then
                    ProcValorImposto txtPedido, IIf(TBProduto!ID_CF = "", 0, TBProduto!ID_CF), IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), 0
                    ProcControleImposto IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), IIf(txtIDfornecedor = "", 0, txtIDfornecedor)
                    If TemIPI = "SIM" Then TBGravar!IPI = IntIPI
                    If TemICMS = "SIM" Then TBGravar!ICMS = IntICMS
                End If
            End If
            
            TBGravar!VlrIPI = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!IPI), 0, TBGravar!IPI)) / 100, "###,##0.00")
            TBGravar!vlrICMS = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!ICMS), 0, TBGravar!ICMS)) / 100, "###,##0.00")
            
            ProcCalculaBCICMSsNecSolic
        End If
    End If
    TBProduto.Close
    
    TBGravar!preco_total = Format((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) + IIf(IsNull(TBGravar!VlrIPI), 0, TBGravar!VlrIPI), "###,##0.00")
    TBGravar!Obs_pedido = IIf(IsNull(TBGravar!Obs), "", TBGravar!Obs)
    TBGravar.Update
    
    '==================================
    Modulo = "Compras/Pedido"
    Evento = "Novo" & IIf(TBGravar!Tipo = "P", "produto", "serviço")
    ID_documento = TBGravar!IDlista
    Documento = "Nº pedido: " & txtPedido
    Documento1 = "Cód. interno: " & TBGravar!Desenho
    ProcGravaEvento
    '==================================
    
    'Atualiza ID do pedido no centro de custo
    Conexao.Execute "Update compras_pedido_lista_custo Set IDpedido = " & txtIDPedido & " where IDlista = " & TBGravar!IDlista
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Responsavel, Data, Percentual, ID_CC, ID_Requisicao from Compras_pedido_lista_custo where idlista = " & TBLISTA!IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            Set TBCodigoDesc = CreateObject("adodb.recordset")
            TBCodigoDesc.Open "select * from Compras_pedido_lista_custo", Conexao, adOpenKeyset, adLockOptimistic
            TBCodigoDesc.AddNew
            TBCodigoDesc!IDpedido = txtIDPedido
            TBCodigoDesc!IDlista = TBGravar!IDlista
            TBCodigoDesc!Responsavel = TBCFOP!Responsavel
            TBCodigoDesc!Data = TBCFOP!Data
            TBCodigoDesc!Percentual = TBCFOP!Percentual
            TBCodigoDesc!ID_CC = TBCFOP!ID_CC
            TBCodigoDesc!ID_Requisicao = TBCFOP!ID_Requisicao
            TBCodigoDesc!valor = (IIf(IsNull(TBCFOP!Percentual), 0, TBCFOP!Percentual) * IIf(IsNull(TBGravar!preco_total), 0, TBGravar!preco_total)) / 100
            TBCodigoDesc.Update
            TBCodigoDesc.Close
            TBCFOP.MoveNext
        Loop
    End If
    TBCFOP.Close
    
    TBGravar.Close
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAlterar_Solicitacao(Necess_PI As Boolean)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_pedido_lista where idlista = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!IDpedido = txtIDPedido
    TBGravar!Quant_Comp = valor
    TBGravar!Quant_Comp_PC = FunCalculaQtdePC(TBGravar!Desenho, valor, True, TBGravar!Unidade_com)
    If TBGravar!prazoreq <> "" Then TBGravar!Prazo = TBGravar!prazoreq
    TBGravar!ValorDesconto = 0
    TBGravar!Status_Item = "AGUARDANDO APROVAÇÃO"
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select codProduto, Pcusto, ID_CF, ID_CFOP, Tipo from projproduto where desenho = '" & Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select PCusto from Projproduto_fornecedor where Codproduto = " & TBProduto!Codproduto & " and idfornecedor = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TBGravar!preco_unitario = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto, "###,##0.0000000000"))
        Else
            TBGravar!preco_unitario = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
        End If
        TBFI.Close
        
        If TBGravar!Remessa = True Then
            valor = Format(FunVerificaQtdeEstoque(TBGravar!Desenho, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
            If valor > 0 Then TBGravar!preco_unitario = Format(Valor_total / valor, "###,##0.0000000000")
        End If
        TBGravar!preco_unitario_desconto = TBGravar!preco_unitario
        
        If IsNull(TBGravar!ID_Requisicao) = False Then
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select RequisicaoTexto from Compras_requisicao where ID_requisicao = " & TBGravar!ID_Requisicao, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then TBGravar!NecessSolici = "Solicitação: " & TBFI!Requisicaotexto & " - Quantidade: " & Format(TBGravar!quant_req, "###,##0.0000") & " - Quantidade PÇ: " & IIf(IsNull(TBGravar!quant_req_PC), 0, TBGravar!quant_req_PC)
            TBFI.Close
        End If
        
        TBGravar!IPI = 0
        TBGravar!ICMS = 0
        If TBProduto!Tipo = "S" Then
            If IsNull(TBProduto!ID_CFOP) = False Then TBGravar!ID_CFOP = TBProduto!ID_CFOP
            
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Simples, presumido, Real, Simples1 from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Regime = 0
                If TBFIltro!Simples = True Then Regime = 1
                If TBFIltro!Presumido = True Then Regime = 2
                If TBFIltro!Real = True Then Regime = 3
                If TBFIltro!Simples1 = True Then Regime = 4
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ISS from Impostos where Regime = " & Regime, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBGravar!ISSQN = IIf(IsNull(TBFI!ISS), 0, TBFI!ISS)
                End If
                TBFI.Close
            End If
            TBGravar!VlrISSQN = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!ISSQN), 0, TBGravar!ISSQN)) / 100, "###,##0.00")
        Else
            If IsNull(TBProduto!ID_CF) = False Then TBGravar!ID_CF = TBProduto!ID_CF
            If TBGravar!Remessa = False Then
                If IsNull(TBProduto!ID_CFOP) = False Then TBGravar!ID_CFOP = TBProduto!ID_CFOP
                If IsNull(TBProduto!ID_CF) = False Then
                    ProcValorImposto txtPedido, IIf(TBProduto!ID_CF = "", 0, TBProduto!ID_CF), IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), 0
                    ProcControleImposto IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), IIf(txtIDfornecedor = "", 0, txtIDfornecedor)
                    If TemIPI = "SIM" Then TBGravar!IPI = IntIPI
                    If TemICMS = "SIM" Then TBGravar!ICMS = IntICMS
                End If
            End If
            
            TBGravar!VlrIPI = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!IPI), 0, TBGravar!IPI)) / 100, "###,##0.00")
            TBGravar!vlrICMS = Format(((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) * IIf(IsNull(TBGravar!ICMS), 0, TBGravar!ICMS)) / 100, "###,##0.00")
            
            ProcCalculaBCICMSsNecSolic
        End If
    End If
    TBProduto.Close

    TBGravar!preco_total = Format((TBGravar!preco_unitario_desconto * TBGravar!Quant_Comp) + IIf(IsNull(TBGravar!VlrIPI), 0, TBGravar!VlrIPI), "###,##0.00")
    TBGravar!Obs_pedido = IIf(IsNull(TBGravar!Obs), "", TBGravar!Obs)
    TBGravar.Update
    
    '==================================
    Modulo = "Compras/Pedido"
    Evento = "Novo" & IIf(TBGravar!Tipo = "P", "produto", "serviço")
    ID_documento = TBGravar!IDlista
    Documento = "Nº pedido: " & txtPedido
    Documento1 = "Cód. interno: " & TBGravar!Desenho
    ProcGravaEvento
    '==================================
    
    'Atualiza ID do pedido no centro de custo
    ValorTotal = IIf(IsNull(TBGravar!preco_total), 0, TBGravar!preco_total)
    NovoValor = Replace(ValorTotal, ",", ".")
    Conexao.Execute "Update compras_pedido_lista_custo Set IDpedido = " & txtIDPedido & ", valor = (ISNULL(Percentual, 0) * " & NovoValor & ") / 100 where IDlista = " & TBGravar!IDlista
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaBCICMSsNecSolic()
On Error GoTo tratar_erro

If IsNull(TBProduto!ID_CFOP) = False Then
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select id_CFOP from tbl_NaturezaOperacao where IDCountCfop = " & IIf(TBProduto!ID_CFOP = "", 0, TBProduto!ID_CFOP), Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select CST_ICMS from tbl_NaturezaOperacao_CST where ID_CFOP = " & IIf(TBProduto!ID_CFOP = "", 0, TBProduto!ID_CFOP), Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If TBFI.RecordCount = 1 Then TBGravar!CST = TBFI!CST_ICMS
    End If
    TBFI.Close
    ProcCalculaBC Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBFIltro!ID_CFOP, 0, (TBGravar!preco_unitario * valor), TBGravar!VlrIPI, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBGravar!CST), "", TBGravar!CST), "T", txtIDfornecedor, txtFornecedor
    TBGravar!BC_ICMS = BC

    If IsNull(TBProduto!ID_CF) = False Then
        If IsNull(TBGravar!CST) = False And TBGravar!Remessa = False Then
            ProcSubstituicaoTributaria txtuf, TBGravar!CST, TBProduto!ID_CF, IIf(txtIDfornecedor = "", 0, txtIDfornecedor), txtFornecedor, TBGravar!preco_unitario, valor, BC, BCST, 0, 0, 0, False, False, 0
            TBGravar!Valor_ICMS_ST = ICMSCST
            If ICMSCST <> 0 Then TBGravar!BC_ICMS_ST = BCICMSCST Else TBGravar!BC_ICMS_ST = 0
        Else
            TBGravar!Valor_ICMS_ST = 0
            TBGravar!BC_ICMS_ST = 0
        End If
    End If
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCamposListaPagina(IndexTexto As Integer)
On Error GoTo tratar_erro

If IndexTexto = 1 Then ListaNecessidade.ListItems.Clear Else Lista_solicitados.ListItems.Clear
lblRegistros(IndexTexto).Caption = "Nº de registros: 0"
lblPaginas(IndexTexto).Caption = "Página: 0 de: 0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
