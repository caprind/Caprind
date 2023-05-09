VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_PI 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Pedido interno"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15390
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
   Icon            =   "frmVendas_PI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15390
   WindowState     =   2  'Maximized
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
      Index           =   0
      Left            =   3660
      TabIndex        =   319
      Top             =   8610
      Width           =   11685
      Begin VB.TextBox txtTotalfrete 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   327
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do frete"
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox txt_TotalIPI 
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
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   326
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do IPI."
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_VlrTotalProd 
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
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   325
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do(s) produto(s)."
         Top             =   360
         Width           =   1245
      End
      Begin VB.TextBox txtTotalservicos 
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
         Height          =   315
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   324
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do(s) serviço(s)"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txttotalproposta 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   10260
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   323
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total da proposta."
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtTotaldesconto 
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
         Height          =   315
         Left            =   2955
         MaxLength       =   50
         TabIndex        =   322
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do desconto."
         Top             =   360
         Width           =   1155
      End
      Begin VB.TextBox txt_ValorNota 
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
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   321
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Subtotal."
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox txt_ICMSs 
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
         Height          =   315
         Left            =   8805
         MaxLength       =   15
         TabIndex        =   320
         Text            =   "0,00"
         ToolTipText     =   "Valor do ICMS substituto."
         Top             =   360
         Width           =   1215
      End
      Begin DrawSuite2022.USButton btnSalvarFrete 
         Height          =   315
         Left            =   6780
         TabIndex        =   379
         ToolTipText     =   "Salvar Valor do frete no(s) item(ns) da lista..."
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         DibPicture      =   "frmVendas_PI.frx":000C
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
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   143
         Left            =   5490
         TabIndex        =   380
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total frete"
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
         Index           =   142
         Left            =   5790
         TabIndex        =   378
         Top             =   150
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "="
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   136
         Left            =   4170
         TabIndex        =   341
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   137
         Left            =   7245
         TabIndex        =   340
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   138
         Left            =   8595
         TabIndex        =   339
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "="
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   139
         Left            =   10080
         TabIndex        =   338
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   1380
         TabIndex        =   337
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   135
         Left            =   2805
         TabIndex        =   336
         Top             =   420
         Width           =   60
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS ST"
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
         Index           =   79
         Left            =   8850
         TabIndex        =   335
         Top             =   150
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pedido"
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
         Index           =   80
         Left            =   10365
         TabIndex        =   334
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total desc."
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
         Index           =   76
         Left            =   3060
         TabIndex        =   333
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal"
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
         Index           =   77
         Left            =   4590
         TabIndex        =   332
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total IPI"
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
         Index           =   78
         Left            =   7635
         TabIndex        =   331
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total serviços"
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
         Index           =   75
         Left            =   1560
         TabIndex        =   330
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total produtos"
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
         Index           =   4
         Left            =   90
         TabIndex        =   329
         Top             =   150
         Width           =   1245
      End
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
      FormWidthDT     =   15510
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15390
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
      Index           =   2
      Left            =   60
      TabIndex        =   202
      Top             =   8610
      Width           =   3615
      Begin VB.TextBox txt_baseICMSs 
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
         Height          =   315
         Left            =   2250
         MaxLength       =   15
         TabIndex        =   205
         Text            =   "0,00"
         ToolTipText     =   "Base de calculo ICMS substituto."
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_VlrICMS 
         Alignment       =   2  'Center
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
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   204
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do ICMS."
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox txt_BaseICMS 
         Alignment       =   2  'Center
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
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   203
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Base de cálculo do ICMS."
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BC ICMS ST"
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
         Index           =   74
         Left            =   2325
         TabIndex        =   268
         Top             =   150
         Width           =   945
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS"
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
         Index           =   73
         Left            =   1260
         TabIndex        =   267
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BC do ICMS"
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
         Index           =   31
         Left            =   165
         TabIndex        =   206
         Top             =   150
         Width           =   945
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   184
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Informações gerais"
      TabPicture(0)   =   "frmVendas_PI.frx":00D3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(2)=   "Frame1(3)"
      Tab(0).Control(3)=   "txtidequipamento"
      Tab(0).Control(4)=   "txtid"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lista"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Dados comerciais"
      TabPicture(1)   =   "frmVendas_PI.frx":00EF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(1)=   "Frame1(4)"
      Tab(1).Control(2)=   "Txt_ID_cobranca"
      Tab(1).Control(3)=   "Txt_ID_entrega"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Lista de produtos"
      TabPicture(2)   =   "frmVendas_PI.frx":010B
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USProgressBar1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SSTab2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "USToolBar3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1(6)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Listprod"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtid_produto"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Lista de serviços"
      TabPicture(3)   =   "frmVendas_PI.frx":0127
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "USToolBar4"
      Tab(3).Control(1)=   "Frame1(13)"
      Tab(3).Control(2)=   "Frame1(8)"
      Tab(3).Control(3)=   "Chk_CFOP_serv"
      Tab(3).Control(4)=   "Chk_obs_faturamento_serv"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Escopo de fornecimento"
      TabPicture(4)   =   "frmVendas_PI.frx":0143
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "USToolBar5"
      Tab(4).Control(1)=   "Frame1(9)"
      Tab(4).ControlCount=   2
      Begin MSComctlLib.ListView Lista 
         Height          =   3600
         Left            =   -74940
         TabIndex        =   35
         Top             =   5790
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   6350
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Proposta"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   13944
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Valor total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Validada"
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.TextBox txtid_produto 
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
         Left            =   2100
         TabIndex        =   223
         Text            =   "0"
         Top             =   5700
         Visible         =   0   'False
         Width           =   465
      End
      Begin MSComctlLib.ListView Listprod 
         Height          =   4290
         Left            =   60
         TabIndex        =   222
         Top             =   4350
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   7567
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
         MousePointer    =   99
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
            Text            =   "Cód. interno"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   7788
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Vlr. unitário"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Desc. (%)"
            Object.Width           =   1499
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
            Text            =   "Vlr. unit. c/ desc."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Vlr. total"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "D"
            Text            =   "Pr. final"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Ped. cliente"
            Object.Width           =   2381
         EndProperty
      End
      Begin VB.CheckBox Chk_obs_faturamento_serv 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -65400
         TabIndex        =   196
         Top             =   2670
         Width           =   195
      End
      Begin VB.CheckBox Chk_CFOP_serv 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -69300
         TabIndex        =   195
         Top             =   1560
         Width           =   195
      End
      Begin VB.TextBox Txt_ID_entrega 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   -71010
         TabIndex        =   194
         Text            =   "0"
         Top             =   8160
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox Txt_ID_cobranca 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   -70650
         TabIndex        =   193
         Text            =   "0"
         Top             =   8160
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   8
         Left            =   -74940
         TabIndex        =   190
         Top             =   9345
         Width           =   15285
         Begin VB.ComboBox cmbOpcao_lista_serv 
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
            ItemData        =   "frmVendas_PI.frx":015F
            Left            =   6840
            List            =   "frmVendas_PI.frx":0169
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   180
            Width           =   1965
         End
         Begin VB.TextBox txtPagIr2 
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
            Left            =   9540
            TabIndex        =   152
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg2 
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
            Left            =   2880
            TabIndex        =   150
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx2 
            Height          =   315
            Left            =   11760
            TabIndex        =   156
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":017E
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
         Begin DrawSuite2022.USButton cmdPagAnt2 
            Height          =   315
            Left            =   11220
            TabIndex        =   155
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":3922
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
         Begin DrawSuite2022.USButton cmdPagIr2 
            Height          =   315
            Left            =   10110
            TabIndex        =   153
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
         Begin DrawSuite2022.USButton cmdPagPrim2 
            Height          =   315
            Left            =   10680
            TabIndex        =   154
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":742B
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
         Begin DrawSuite2022.USButton cmdPagUlt2 
            Height          =   315
            Left            =   12300
            TabIndex        =   157
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":B51A
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
            Index           =   49
            Left            =   3510
            TabIndex        =   245
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   5520
            TabIndex        =   216
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   2190
            TabIndex        =   211
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblPaginas2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   192
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   191
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   6
         Left            =   60
         TabIndex        =   187
         Top             =   9375
         Width           =   15285
         Begin VB.ComboBox cmbOpcao_lista_prod 
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
            ItemData        =   "frmVendas_PI.frx":EDA6
            Left            =   6840
            List            =   "frmVendas_PI.frx":EDB0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   180
            Width           =   1965
         End
         Begin VB.TextBox txtNreg1 
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
            Left            =   2880
            TabIndex        =   99
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr1 
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
            Left            =   9540
            TabIndex        =   101
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx1 
            Height          =   315
            Left            =   11760
            TabIndex        =   105
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":EDC5
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
         Begin DrawSuite2022.USButton cmdPagAnt1 
            Height          =   315
            Left            =   11220
            TabIndex        =   104
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":12569
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
         Begin DrawSuite2022.USButton cmdPagIr1 
            Height          =   315
            Left            =   10110
            TabIndex        =   102
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
         Begin DrawSuite2022.USButton cmdPagPrim1 
            Height          =   315
            Left            =   10680
            TabIndex        =   103
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":16072
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
         Begin DrawSuite2022.USButton cmdPagUlt1 
            Height          =   315
            Left            =   12300
            TabIndex        =   106
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":1A161
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
            Index           =   47
            Left            =   3510
            TabIndex        =   244
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   5520
            TabIndex        =   215
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   2190
            TabIndex        =   210
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   189
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   188
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtid 
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
         Height          =   345
         Left            =   -72750
         Locked          =   -1  'True
         TabIndex        =   183
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "ID da proposta"
         Top             =   7560
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtidequipamento 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72360
         MaxLength       =   50
         MousePointer    =   99  'Custom
         TabIndex        =   182
         ToolTipText     =   "Descrição."
         Top             =   7560
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   3
         Left            =   -74940
         TabIndex        =   179
         Top             =   9360
         Width           =   15285
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
            ItemData        =   "frmVendas_PI.frx":1D9ED
            Left            =   6810
            List            =   "frmVendas_PI.frx":1D9F7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   180
            Width           =   1965
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
            Left            =   2880
            TabIndex        =   36
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
            Left            =   9540
            TabIndex        =   38
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   42
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":1DA0E
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
            Left            =   11220
            TabIndex        =   41
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":211B2
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
            Left            =   10110
            TabIndex        =   39
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
            Left            =   10680
            TabIndex        =   40
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":24CBB
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
            Left            =   12300
            TabIndex        =   43
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":28DAA
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
            Index           =   48
            Left            =   3510
            TabIndex        =   243
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   5520
            TabIndex        =   214
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   2190
            TabIndex        =   209
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   181
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   180
            Top             =   240
            Width           =   1095
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   1005
         Left            =   -74940
         TabIndex        =   173
         Top             =   360
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   1773
         ButtonCount     =   19
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
         ButtonLeft2     =   37
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   75
         ButtonTop3      =   2
         ButtonWidth3    =   38
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
         ButtonLeft4     =   115
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
         ButtonLeft5     =   168
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
         ButtonLeft6     =   217
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Copiar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Copiar (F7)"
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
         ButtonLeft7     =   265
         ButtonTop7      =   2
         ButtonWidth7    =   39
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Necessidades"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Check lista de necessidades de compras"
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
         ButtonLeft8     =   306
         ButtonTop8      =   2
         ButtonWidth8    =   73
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Revisar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Revisar (F8)"
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
         ButtonLeft9     =   381
         ButtonTop9      =   2
         ButtonWidth9    =   44
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Emitir PI"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Emitir pedido interno (F9)"
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
         ButtonLeft10    =   427
         ButtonTop10     =   2
         ButtonWidth10   =   47
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Cancelar PI"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Cancelar pedido interno (F10)"
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
         ButtonLeft11    =   476
         ButtonTop11     =   2
         ButtonWidth11   =   63
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Impostos"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Visualizar impostos."
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
         ButtonLeft12    =   541
         ButtonTop12     =   2
         ButtonWidth12   =   52
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Status"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Alterar status (F11)"
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
         ButtonLeft13    =   595
         ButtonTop13     =   2
         ButtonWidth13   =   39
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Validação"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Validar/cancelar validação (F12)"
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
         ButtonLeft14    =   636
         ButtonTop14     =   2
         ButtonWidth14   =   53
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonCaption15 =   "Importar"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Importar pedido de compra (industrialização) do excel (F12)"
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   691
         ButtonTop15     =   2
         ButtonWidth15   =   50
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonCaption16 =   "Atualizar"
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonToolTipText16=   "Utilizado pelo administrador do sistema."
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft16    =   743
         ButtonTop16     =   2
         ButtonWidth16   =   50
         ButtonHeight16  =   21
         ButtonUseMaskColor16=   0   'False
         ButtonEnabled17 =   0   'False
         ButtonIconSize17=   32
         ButtonAlignment17=   2
         ButtonType17    =   1
         ButtonStyle17   =   -1
         BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState17   =   -1
         ButtonLeft17    =   795
         ButtonTop17     =   4
         ButtonWidth17   =   2
         ButtonHeight17  =   56
         ButtonCaption18 =   "Ajuda"
         ButtonEnabled18 =   0   'False
         ButtonIconSize18=   32
         ButtonToolTipText18=   "Ajuda (F1)"
         ButtonKey18     =   "17"
         ButtonAlignment18=   2
         BeginProperty ButtonFont18 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft18    =   799
         ButtonTop18     =   2
         ButtonWidth18   =   36
         ButtonHeight18  =   21
         ButtonUseMaskColor18=   0   'False
         ButtonCaption19 =   "Sair"
         ButtonEnabled19 =   0   'False
         ButtonIconSize19=   32
         ButtonToolTipText19=   "Sair (Esc)"
         ButtonKey19     =   "18"
         ButtonAlignment19=   2
         BeginProperty ButtonFont19 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft19    =   837
         ButtonTop19     =   2
         ButtonWidth19   =   26
         ButtonHeight19  =   21
         ButtonUseMaskColor19=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12270
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_PI.frx":2C636
            Count           =   1
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   4440
         Index           =   1
         Left            =   -74945
         TabIndex        =   168
         Top             =   1350
         Width           =   15285
         Begin VB.TextBox txtDatavendas 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   13800
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   411
            ToolTipText     =   "Data da venda."
            Top             =   2715
            Width           =   1230
         End
         Begin VB.TextBox txtIE 
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
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   21
            TabIndex        =   407
            ToolTipText     =   "Número do fax."
            Top             =   1560
            Width           =   1365
         End
         Begin VB.TextBox txtCEP 
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
            MaxLength       =   30
            TabIndex        =   404
            ToolTipText     =   "Complemento."
            Top             =   2130
            Width           =   915
         End
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   21
            TabIndex        =   381
            ToolTipText     =   "Número do fax."
            Top             =   1560
            Width           =   1455
         End
         Begin DrawSuite2022.USButton cmdstatus 
            Height          =   315
            Left            =   6120
            TabIndex        =   343
            Top             =   960
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":38278
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
         Begin VB.TextBox txtRespValidacao 
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
            Left            =   12120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   375
            Width           =   2910
         End
         Begin VB.TextBox txtDtValidacao 
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
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   375
            Width           =   2025
         End
         Begin VB.ComboBox cmbCidade 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmVendas_PI.frx":5637D
            Left            =   1120
            List            =   "frmVendas_PI.frx":5637F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Cidade."
            Top             =   2715
            Width           =   3665
         End
         Begin VB.TextBox txtBairro 
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
            Left            =   11490
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   22
            ToolTipText     =   "Bairro."
            Top             =   2130
            Width           =   3540
         End
         Begin VB.ComboBox cmbTipo_bairro 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmVendas_PI.frx":56381
            Left            =   10350
            List            =   "frmVendas_PI.frx":563C1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            ToolTipText     =   "Tipo do bairro."
            Top             =   2130
            Width           =   1140
         End
         Begin VB.ComboBox cmbTipo_endereco 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmVendas_PI.frx":56471
            Left            =   1110
            List            =   "frmVendas_PI.frx":564B1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Tipo do endereço."
            Top             =   2130
            Width           =   1140
         End
         Begin VB.TextBox txtendereco 
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
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   18
            ToolTipText     =   "Endereço."
            Top             =   2130
            Width           =   6180
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   8445
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   19
            ToolTipText     =   "Número."
            Top             =   2130
            Width           =   720
         End
         Begin VB.TextBox txtComplemento 
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
            Left            =   9180
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   20
            ToolTipText     =   "Complemento."
            Top             =   2130
            Width           =   1155
         End
         Begin VB.TextBox txtregiao 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   12000
            Locked          =   -1  'True
            MaxLength       =   180
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Região do vendedor externo."
            Top             =   3300
            Width           =   3015
         End
         Begin VB.TextBox txtreferente 
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
            Left            =   9285
            MaxLength       =   180
            TabIndex        =   27
            ToolTipText     =   "Descrição da referência."
            Top             =   2710
            Width           =   4495
         End
         Begin VB.TextBox txt_datamodificado 
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
            Left            =   1730
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão/cancelamento/perda."
            Top             =   960
            Width           =   1710
         End
         Begin VB.TextBox txt_observacoes 
            BackColor       =   &H00FFFFFF&
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
            Height          =   405
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            ToolTipText     =   "Observações gerais."
            Top             =   3900
            Width           =   14835
         End
         Begin VB.TextBox txtresponsavel 
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
            Left            =   7155
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   375
            Width           =   2910
         End
         Begin VB.TextBox txtCotacao 
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Número da proposta comercial."
            Top             =   955
            Width           =   1035
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
            Height          =   315
            Left            =   1120
            MaxLength       =   60
            TabIndex        =   25
            ToolTipText     =   "Cidade."
            Top             =   2710
            Visible         =   0   'False
            Width           =   3665
         End
         Begin VB.TextBox txtCliente 
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
            Left            =   7080
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   10
            ToolTipText     =   "Nome do cliente."
            Top             =   955
            Width           =   6960
         End
         Begin VB.TextBox txtdepartamento 
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
            Left            =   5910
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   13
            ToolTipText     =   "Departamento do contato."
            Top             =   1560
            Width           =   3180
         End
         Begin VB.TextBox txtVE 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   6120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Vendedor externo."
            Top             =   3300
            Width           =   555
         End
         Begin VB.TextBox txtVI 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Vendedor interno."
            Top             =   3300
            Width           =   585
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
            Height          =   315
            Left            =   9105
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   14
            ToolTipText     =   "Número do telefone."
            Top             =   1560
            Width           =   1185
         End
         Begin VB.TextBox txtFax 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   10305
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   15
            ToolTipText     =   "Número do fax."
            Top             =   1560
            Width           =   1185
         End
         Begin VB.TextBox txtIDCliente 
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
            Left            =   6540
            TabIndex        =   9
            ToolTipText     =   "Código do cliente."
            Top             =   955
            Width           =   530
         End
         Begin VB.TextBox txtRef 
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
            Left            =   4800
            MaxLength       =   100
            TabIndex        =   26
            ToolTipText     =   "Número da referência."
            Top             =   2710
            Width           =   4470
         End
         Begin VB.TextBox txtRemetente 
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
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            ToolTipText     =   "Contato do cliente."
            Top             =   1560
            Width           =   2565
         End
         Begin VB.TextBox txtvend_Int 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   780
            Locked          =   -1  'True
            MaxLength       =   180
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Vendedor interno."
            Top             =   3300
            Width           =   4890
         End
         Begin VB.TextBox txtVend_Ext 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   6690
            Locked          =   -1  'True
            MaxLength       =   180
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Vendedor externo."
            Top             =   3300
            Width           =   4860
         End
         Begin VB.TextBox txtstatus 
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
            Height          =   315
            Left            =   3450
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtrevisao 
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
            Height          =   315
            Left            =   1225
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Revisão da proposta."
            Top             =   960
            Width           =   485
         End
         Begin VB.ComboBox txttipocliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
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
            ItemData        =   "frmVendas_PI.frx":56548
            Left            =   14055
            List            =   "frmVendas_PI.frx":56558
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Tipo do cliente."
            Top             =   955
            Width           =   630
         End
         Begin VB.ComboBox txtuf 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmVendas_PI.frx":5656C
            Left            =   180
            List            =   "frmVendas_PI.frx":5656E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "UF."
            Top             =   2710
            Width           =   930
         End
         Begin VB.TextBox txtEmail 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   11490
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   16
            ToolTipText     =   "E-mail."
            Top             =   1560
            Width           =   3540
         End
         Begin VB.ComboBox Cmb_empresa 
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
            ItemData        =   "frmVendas_PI.frx":56570
            Left            =   180
            List            =   "frmVendas_PI.frx":56572
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   375
            Width           =   5715
         End
         Begin MSComCtl2.DTPicker txtDatavendas_PI 
            Height          =   315
            Left            =   13800
            TabIndex        =   28
            ToolTipText     =   "Data da venda."
            Top             =   2710
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   103284739
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txt_dataelaborado 
            Height          =   315
            Left            =   5910
            TabIndex        =   1
            ToolTipText     =   "Data de emissão."
            Top             =   375
            Width           =   1230
            _ExtentX        =   2170
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
            Format          =   103284739
            CurrentDate     =   39057
         End
         Begin DrawSuite2022.USButton cmdadicionarcliente 
            Height          =   315
            Left            =   14700
            TabIndex        =   344
            Top             =   960
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":56574
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
         Begin DrawSuite2022.USButton cmdcontato 
            Height          =   315
            Left            =   5580
            TabIndex        =   345
            Top             =   1560
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":74679
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
         Begin DrawSuite2022.USButton cmdVendedor_Interno 
            Height          =   315
            Left            =   5700
            TabIndex        =   346
            Top             =   3300
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":9277E
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
         Begin DrawSuite2022.USButton Cmdvendedor 
            Height          =   315
            Left            =   11550
            TabIndex        =   347
            Top             =   3300
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":B0883
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
            Caption         =   "Inscrição estadual"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   145
            Left            =   1680
            TabIndex        =   406
            Top             =   1350
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Index           =   10
            Left            =   510
            TabIndex        =   405
            Top             =   1950
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ"
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
            Index           =   144
            Left            =   690
            TabIndex        =   382
            Top             =   1350
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Região"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   13260
            TabIndex        =   266
            Top             =   3090
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor externo ( Representante )"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   71
            Left            =   7800
            TabIndex        =   265
            Top             =   3090
            Width           =   2640
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição da referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   70
            Left            =   10685
            TabIndex        =   264
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Número da referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   69
            Left            =   6255
            TabIndex        =   263
            Top             =   2520
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   68
            Left            =   2705
            TabIndex        =   262
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   67
            Left            =   13050
            TabIndex        =   261
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   66
            Left            =   10770
            TabIndex        =   260
            Top             =   1920
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   65
            Left            =   9270
            TabIndex        =   259
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   8520
            TabIndex        =   258
            Top             =   1920
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   63
            Left            =   5010
            TabIndex        =   257
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   62
            Left            =   12750
            TabIndex        =   256
            Top             =   1350
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   10755
            TabIndex        =   255
            Top             =   1350
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   60
            Left            =   9375
            TabIndex        =   254
            Top             =   1350
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   59
            Left            =   6990
            TabIndex        =   253
            Top             =   1350
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   58
            Left            =   14220
            TabIndex        =   252
            Top             =   750
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   57
            Left            =   10313
            TabIndex        =   251
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   56
            Left            =   4500
            TabIndex        =   250
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   1280
            TabIndex        =   249
            Top             =   750
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável validação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   12758
            TabIndex        =   248
            Top             =   180
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora validação"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   10365
            TabIndex        =   247
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   8153
            TabIndex        =   246
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Revisada em"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   2128
            TabIndex        =   207
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data da venda"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   13883
            TabIndex        =   200
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido int."
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
            Index           =   15
            Left            =   315
            TabIndex        =   199
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label1 
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
            Index           =   14
            Left            =   2760
            TabIndex        =   198
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. de emissão"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   5993
            TabIndex        =   197
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   548
            TabIndex        =   186
            Top             =   2520
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   1530
            TabIndex        =   185
            Top             =   1920
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   7125
            TabIndex        =   171
            Top             =   3690
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor interno"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   2603
            TabIndex        =   170
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Contato"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3990
            TabIndex        =   169
            Top             =   1350
            Width           =   585
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
         Height          =   8640
         Index           =   9
         Left            =   -74940
         TabIndex        =   165
         Top             =   1350
         Width           =   15285
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
            Height          =   6975
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   158
            ToolTipText     =   "Escopo de fornecimento."
            Top             =   210
            Width           =   14805
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
         Height          =   7185
         Index           =   13
         Left            =   -74940
         TabIndex        =   160
         Top             =   1350
         Width           =   15285
         Begin MSComctlLib.ListView Listaservicos 
            Height          =   3870
            Left            =   180
            TabIndex        =   149
            Top             =   2910
            Width           =   14805
            _ExtentX        =   26114
            _ExtentY        =   6826
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
            MousePointer    =   99
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
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   7788
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Vlr. unitário"
               Object.Width           =   1676
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Desc. (%)"
               Object.Width           =   1499
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
               Text            =   "Vlr. unit. c/ desc."
               Object.Width           =   2205
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Vlr. total"
               Object.Width           =   1676
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "T"
               Text            =   "Ped. cliente"
               Object.Width           =   1587
            EndProperty
         End
         Begin VB.TextBox txtid_servico 
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
            Left            =   2460
            TabIndex        =   164
            Text            =   "0"
            Top             =   4170
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   3075
            Index           =   11
            Left            =   30
            TabIndex        =   161
            Top             =   150
            Width           =   15135
            Begin VB.CheckBox Chk_utiliza_mat_consignado_serv 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Utiliza material consignado?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12510
               TabIndex        =   403
               Top             =   1920
               Width           =   2445
            End
            Begin VB.ComboBox Cmb_cidade_servico 
               Appearance      =   0  'Flat
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
               ItemData        =   "frmVendas_PI.frx":CE988
               Left            =   180
               List            =   "frmVendas_PI.frx":CE98A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   135
               ToolTipText     =   "Cidade onde foi executado o serviço."
               Top             =   2355
               Width           =   3090
            End
            Begin VB.TextBox txtdesccomservico 
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
               Height          =   675
               Left            =   180
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   129
               ToolTipText     =   "Descrição comercial."
               Top             =   1395
               Width           =   6345
            End
            Begin VB.CommandButton Cmd_importar_PC_serv 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11310
               Picture         =   "frmVendas_PI.frx":CE98C
               Style           =   1  'Graphical
               TabIndex        =   124
               ToolTipText     =   "Localizar caminho do pedido de compra."
               Top             =   840
               Width           =   315
            End
            Begin VB.CheckBox Chk_prazo_serv 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12390
               TabIndex        =   213
               Top             =   630
               Width           =   195
            End
            Begin VB.CheckBox Chk_PC_serv 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9810
               TabIndex        =   212
               Top             =   630
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.TextBox txtpcclienteserv 
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
               Left            =   9810
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   123
               TabStop         =   0   'False
               ToolTipText     =   "Pedido do cliente."
               Top             =   840
               Width           =   1485
            End
            Begin VB.CommandButton Cmd_limpar_CFOP_serv 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11220
               Picture         =   "frmVendas_PI.frx":CEA8E
               Style           =   1  'Graphical
               TabIndex        =   118
               ToolTipText     =   "Limpar CFOP."
               Top             =   285
               Width           =   315
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   3300
               Picture         =   "frmVendas_PI.frx":CEBCC
               Style           =   1  'Graphical
               TabIndex        =   111
               ToolTipText     =   "Visualizar arquivo."
               Top             =   285
               Width           =   315
            End
            Begin VB.TextBox Txt_observacoes_fat_serv 
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
               Height          =   675
               Left            =   9480
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   131
               ToolTipText     =   "Observações de faturamento."
               Top             =   1395
               Width           =   2925
            End
            Begin VB.CheckBox Chk_antecipacao_serv 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Antecipação de faturamento?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12510
               TabIndex        =   133
               Top             =   1455
               Width           =   2445
            End
            Begin VB.CheckBox Chk_faturamento_parcial_serv 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Faturamento parcial?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12510
               TabIndex        =   134
               Top             =   1695
               Width           =   1815
            End
            Begin VB.ComboBox Cmb_un_com_serv 
               Appearance      =   0  'Flat
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
               Left            =   4005
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   137
               ToolTipText     =   "Unidade comercial."
               Top             =   2355
               Width           =   735
            End
            Begin VB.TextBox Txt_ID_CFOP_serv 
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
               Left            =   5580
               Locked          =   -1  'True
               TabIndex        =   114
               TabStop         =   0   'False
               ToolTipText     =   "ID da CFOP."
               Top             =   285
               Width           =   525
            End
            Begin VB.TextBox Txt_natureza_operacao_serv 
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
               Left            =   7200
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   116
               TabStop         =   0   'False
               ToolTipText     =   "Descrição da natureza da operação."
               Top             =   285
               Width           =   3675
            End
            Begin VB.CommandButton Cmd_localizar_CFOP_serv 
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
               Left            =   10890
               Picture         =   "frmVendas_PI.frx":CF18E
               Style           =   1  'Graphical
               TabIndex        =   117
               ToolTipText     =   "Localizar CFOP."
               Top             =   285
               Width           =   315
            End
            Begin VB.TextBox txtObs_serv 
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
               Height          =   675
               Left            =   6540
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   130
               ToolTipText     =   "Observações."
               Top             =   1395
               Width           =   2925
            End
            Begin VB.CheckBox Chk_valor_desc2 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   8220
               TabIndex        =   142
               Top             =   2415
               Width           =   225
            End
            Begin VB.CheckBox Chk_desc2 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6915
               TabIndex        =   140
               Top             =   2415
               Width           =   225
            End
            Begin VB.TextBox txtComissaoServ 
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
               Left            =   14070
               MaxLength       =   15
               TabIndex        =   148
               ToolTipText     =   "Comissão do vendedor."
               Top             =   2355
               Width           =   885
            End
            Begin VB.TextBox Txt_analise1 
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
               Left            =   13830
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   127
               TabStop         =   0   'False
               ToolTipText     =   "Análise crítica."
               Top             =   840
               Width           =   795
            End
            Begin VB.CommandButton Cmd_analise1 
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
               Left            =   14640
               Picture         =   "frmVendas_PI.frx":CF290
               Style           =   1  'Graphical
               TabIndex        =   128
               ToolTipText     =   "Abrir lista de análises aprovadas."
               Top             =   840
               Width           =   315
            End
            Begin VB.CheckBox Chk_servico_executado_cliente 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Executado no cliente?"
               Height          =   195
               Left            =   12510
               TabIndex        =   132
               Top             =   1215
               Width           =   1875
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo serviço"
               ForeColor       =   &H00000000&
               Height          =   555
               Index           =   7
               Left            =   11640
               TabIndex        =   167
               Top             =   60
               Width           =   3345
               Begin VB.CheckBox OPTnovoservicoman 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. manual ?"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   1860
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CheckBox optnovoservico 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Cód. automático ?"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   180
                  TabIndex        =   119
                  Top             =   240
                  Width           =   1605
               End
            End
            Begin VB.TextBox txtvalorunitariodesc2 
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
               Height          =   315
               Left            =   9630
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   144
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   2355
               Width           =   1380
            End
            Begin VB.TextBox txtvalordesconto2 
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
               Height          =   315
               Left            =   8460
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   143
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   2355
               Width           =   1155
            End
            Begin VB.TextBox txtiss 
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
               Left            =   11025
               TabIndex        =   145
               ToolTipText     =   "Porcentagem do ISSQN."
               Top             =   2355
               Width           =   930
            End
            Begin VB.TextBox txtvlrISS 
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
               Height          =   315
               Left            =   11970
               Locked          =   -1  'True
               TabIndex        =   146
               TabStop         =   0   'False
               ToolTipText     =   "Valor do ISSQN."
               Top             =   2355
               Width           =   995
            End
            Begin VB.TextBox txtdesconto2 
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
               Left            =   7155
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   141
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   2355
               Width           =   960
            End
            Begin VB.TextBox txtqtservico 
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
               Height          =   315
               Left            =   4740
               MaxLength       =   50
               TabIndex        =   138
               ToolTipText     =   "Quantidade."
               Top             =   2355
               Width           =   930
            End
            Begin VB.TextBox txtvlrunitservico 
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
               Height          =   315
               Left            =   5685
               MaxLength       =   50
               TabIndex        =   139
               ToolTipText     =   "Valor unitário."
               Top             =   2355
               Width           =   1125
            End
            Begin VB.TextBox txtvlrtotalservico 
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
               Height          =   315
               Left            =   12970
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   147
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   2355
               Width           =   1085
            End
            Begin VB.ComboBox txtunservico 
               Appearance      =   0  'Flat
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
               Left            =   3270
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   136
               ToolTipText     =   "Unidade de estoque."
               Top             =   2355
               Width           =   735
            End
            Begin VB.CommandButton cmdfiltrar_serv 
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
               Left            =   2640
               Picture         =   "frmVendas_PI.frx":CF372
               Style           =   1  'Graphical
               TabIndex        =   109
               ToolTipText     =   "Filtrar por código interno."
               Top             =   285
               Width           =   315
            End
            Begin VB.ComboBox cmbreferencia_serv 
               Appearance      =   0  'Flat
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
               Left            =   3690
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   112
               ToolTipText     =   "Código de referência."
               Top             =   285
               Width           =   1890
            End
            Begin VB.ComboBox cmbfamiliaservico 
               Appearance      =   0  'Flat
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
               Left            =   5520
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   122
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   840
               Width           =   4275
            End
            Begin VB.TextBox txtdescservico 
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
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   121
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   840
               Width           =   5325
            End
            Begin VB.CommandButton cmdlistaservicos 
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
               Left            =   2970
               Picture         =   "frmVendas_PI.frx":CF78D
               Style           =   1  'Graphical
               TabIndex        =   110
               ToolTipText     =   "Localizar seviços (F2)"
               Top             =   285
               Width           =   315
            End
            Begin VB.TextBox txtRev_serv 
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
               Left            =   2085
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   108
               TabStop         =   0   'False
               Text            =   "0"
               ToolTipText     =   "Revisão."
               Top             =   285
               Width           =   525
            End
            Begin VB.TextBox txtcodservico 
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
               Height          =   315
               Left            =   180
               MaxLength       =   50
               TabIndex        =   107
               ToolTipText     =   "Código interno."
               Top             =   285
               Width           =   1890
            End
            Begin VB.TextBox Txt_CFOP_serv 
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
               Left            =   6120
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   115
               TabStop         =   0   'False
               ToolTipText     =   "Natureza da operação."
               Top             =   285
               Width           =   1065
            End
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
               Height          =   315
               Left            =   3690
               MaxLength       =   50
               TabIndex        =   113
               ToolTipText     =   "Código de referência."
               Top             =   285
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.TextBox txtPrazo_Servico 
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
               Left            =   12390
               MaxLength       =   4
               TabIndex        =   125
               ToolTipText     =   "Prazo de entrega em dias."
               Top             =   840
               Width           =   1425
            End
            Begin MSMask.MaskEdBox mskprazoservico 
               Height          =   315
               Left            =   12390
               TabIndex        =   126
               ToolTipText     =   "Prazo de entrega."
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "ISSQN (%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   134
               Left            =   11093
               TabIndex        =   318
               Top             =   2160
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor ISSQN"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   133
               Left            =   12032
               TabIndex        =   317
               Top             =   2160
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Comis.(%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   132
               Left            =   14130
               TabIndex        =   316
               Top             =   2160
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   131
               Left            =   13145
               TabIndex        =   315
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor unit. c/ desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   130
               Left            =   9645
               TabIndex        =   314
               Top             =   2160
               Width           =   1350
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   129
               Left            =   8527
               TabIndex        =   313
               Top             =   2160
               Width           =   1020
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Desc. (%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   128
               Left            =   7268
               TabIndex        =   312
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor unitário"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   127
               Left            =   5775
               TabIndex        =   311
               Top             =   2160
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   126
               Left            =   4995
               TabIndex        =   310
               Top             =   2160
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   125
               Left            =   4050
               TabIndex        =   309
               Top             =   2160
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   124
               Left            =   3345
               TabIndex        =   308
               Top             =   2160
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Observações para faturamento"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   123
               Left            =   9810
               TabIndex        =   307
               Top             =   1185
               Width           =   2265
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Observações"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   122
               Left            =   7530
               TabIndex        =   306
               Top             =   1185
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "ACC"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   121
               Left            =   14070
               TabIndex        =   305
               Top             =   630
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido do cliente"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   120
               Left            =   10065
               TabIndex        =   304
               Top             =   630
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Família"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   119
               Left            =   7417
               TabIndex        =   303
               Top             =   630
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Natureza da operação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   117
               Left            =   8235
               TabIndex        =   301
               Top             =   75
               Width           =   1605
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código de referência"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   116
               Left            =   3885
               TabIndex        =   300
               Top             =   75
               Width           =   1500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   115
               Left            =   5880
               TabIndex        =   299
               Top             =   75
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CFOP"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   114
               Left            =   6450
               TabIndex        =   298
               Top             =   75
               Width           =   405
            End
            Begin VB.Image ImgCalendario1 
               Height          =   360
               Left            =   13485
               Picture         =   "frmVendas_PI.frx":CF88F
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   810
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo (dias)"
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
               Index           =   33
               Left            =   12645
               TabIndex        =   208
               Top             =   630
               Width           =   1020
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Rev."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   28
               Left            =   2190
               TabIndex        =   201
               Top             =   75
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição comercial"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   22
               Left            =   2655
               TabIndex        =   176
               Top             =   1185
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Cidade onde será executado o serviço"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   23
               Left            =   315
               TabIndex        =   166
               Top             =   2160
               Width           =   2760
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   21
               Left            =   2497
               TabIndex        =   163
               Top             =   630
               Width           =   690
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
               Index           =   24
               Left            =   510
               TabIndex        =   162
               Top             =   75
               Width           =   1230
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados gerais"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8655
         Index           =   4
         Left            =   -74940
         TabIndex        =   159
         Top             =   1350
         Width           =   15285
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Informações para o frete | Contrato | Moeda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1545
            Index           =   14
            Left            =   180
            TabIndex        =   383
            Top             =   4560
            Width           =   14805
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
               Left            =   12945
               MaxLength       =   60
               TabIndex        =   400
               ToolTipText     =   "Valor da moeda."
               Top             =   1020
               Width           =   1515
            End
            Begin VB.ComboBox cmbMoeda 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmVendas_PI.frx":CFD12
               Left            =   12960
               List            =   "frmVendas_PI.frx":CFD14
               Style           =   2  'Dropdown List
               TabIndex        =   399
               ToolTipText     =   "Moeda."
               Top             =   450
               Width           =   1515
            End
            Begin VB.ComboBox txtAnalize 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmVendas_PI.frx":CFD16
               Left            =   10545
               List            =   "frmVendas_PI.frx":CFD20
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   397
               Top             =   1020
               Width           =   1845
            End
            Begin VB.TextBox txtTipoTransp 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   1
               Left            =   210
               TabIndex        =   396
               Top             =   1020
               Width           =   1215
            End
            Begin VB.TextBox txtTipoTransp 
               Alignment       =   2  'Center
               Height          =   315
               Index           =   0
               Left            =   210
               TabIndex        =   395
               Top             =   450
               Width           =   1215
            End
            Begin VB.TextBox txtTransportadora 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1950
               TabIndex        =   386
               Top             =   450
               Width           =   6585
            End
            Begin VB.TextBox txtRedespacho 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1950
               TabIndex        =   385
               Top             =   1020
               Width           =   6585
            End
            Begin VB.ComboBox cmb_Tipo_Frete 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmVendas_PI.frx":CFD2E
               Left            =   10530
               List            =   "frmVendas_PI.frx":CFD38
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   384
               ToolTipText     =   "Tipo do frete"
               Top             =   450
               Width           =   1875
            End
            Begin DrawSuite2022.USButton btnTransportadora 
               Height          =   315
               Left            =   8550
               TabIndex        =   387
               Top             =   450
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":CFD60
               Caption         =   " Transportadora"
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
               PicSize         =   1
               ShowFocusRect   =   0   'False
               Theme           =   4
            End
            Begin DrawSuite2022.USButton BtnRedespacho 
               Height          =   315
               Left            =   8550
               TabIndex        =   388
               Top             =   1020
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":D57C5
               Caption         =   " Redespacho"
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
               PicAlign        =   3
               ShowFocusRect   =   0   'False
               Theme           =   4
            End
            Begin VB.TextBox txtidTransportadora 
               Height          =   285
               Left            =   1440
               TabIndex        =   389
               Text            =   "Text1"
               Top             =   1860
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor moeda"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   44
               Left            =   13020
               TabIndex        =   402
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Moeda"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   37
               Left            =   13455
               TabIndex        =   401
               Top             =   240
               Width           =   480
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Contrato"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   103
               Left            =   10635
               TabIndex        =   398
               Top             =   810
               Width           =   1665
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Nome"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   13
               Left            =   5040
               TabIndex        =   394
               Top             =   240
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   46
               Left            =   660
               TabIndex        =   393
               Top             =   240
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Redespacho"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   4800
               TabIndex        =   392
               Top             =   810
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   615
               TabIndex        =   391
               Top             =   810
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Frete por conta"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   45
               Left            =   10830
               TabIndex        =   390
               Top             =   240
               Width           =   1125
            End
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
            Height          =   585
            Left            =   7860
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   328
            ToolTipText     =   "Observações."
            Top             =   450
            Width           =   7155
         End
         Begin VB.TextBox txtcalculos 
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
            Height          =   615
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   218
            TabStop         =   0   'False
            ToolTipText     =   "Desenhos e cálculos."
            Top             =   1320
            Width           =   7095
         End
         Begin VB.ComboBox txtlocal_cobranca 
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
            ItemData        =   "frmVendas_PI.frx":DB22A
            Left            =   7860
            List            =   "frmVendas_PI.frx":DB22C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   51
            ToolTipText     =   "Local de cobrança."
            Top             =   4020
            Width           =   6795
         End
         Begin VB.ComboBox txtlocal_entrega 
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
            ItemData        =   "frmVendas_PI.frx":DB22E
            Left            =   180
            List            =   "frmVendas_PI.frx":DB230
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            ToolTipText     =   "Local de entrega."
            Top             =   4020
            Width           =   7095
         End
         Begin VB.TextBox txtimpostos 
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
            Height          =   615
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Impostos."
            Top             =   2205
            Width           =   7095
         End
         Begin VB.TextBox txtValidade 
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
            Height          =   615
            Left            =   7860
            Locked          =   -1  'True
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Prazo de validade."
            Top             =   3105
            Width           =   6795
         End
         Begin VB.TextBox txtReajuste 
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
            Height          =   615
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Reajuste dos preços."
            Top             =   3105
            Width           =   7095
         End
         Begin VB.TextBox txtCondicoes 
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
            Height          =   615
            Left            =   180
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Condições de pagamento."
            Top             =   405
            Width           =   7095
         End
         Begin VB.TextBox txttransporte 
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
            Height          =   645
            Left            =   7860
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Transporte."
            Top             =   1305
            Width           =   6795
         End
         Begin VB.TextBox txtgarantia 
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
            Height          =   615
            Left            =   7860
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Garantia."
            Top             =   2205
            Width           =   6795
         End
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   0
            Left            =   7290
            TabIndex        =   348
            Top             =   420
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":DB232
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
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   1
            Left            =   7290
            TabIndex        =   349
            Top             =   1320
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":F9337
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
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   2
            Left            =   14670
            TabIndex        =   350
            Top             =   1320
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":11743C
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
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   3
            Left            =   7290
            TabIndex        =   351
            Top             =   2220
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":135541
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
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   4
            Left            =   14670
            TabIndex        =   352
            Top             =   2190
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":153646
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
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   5
            Left            =   7290
            TabIndex        =   353
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":17174B
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
         Begin DrawSuite2022.USButton cmdBotao_DadosComerciais 
            Height          =   615
            Index           =   6
            Left            =   14670
            TabIndex        =   354
            Top             =   3090
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   1085
            DibPicture      =   "frmVendas_PI.frx":18F850
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
         Begin DrawSuite2022.USButton cmdlocalentrega 
            Height          =   315
            Left            =   7290
            TabIndex        =   355
            Top             =   4020
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":1AD955
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
         Begin DrawSuite2022.USButton cmdlocalcobranca 
            Height          =   315
            Left            =   14670
            TabIndex        =   356
            Top             =   4020
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmVendas_PI.frx":1B0FA5
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
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Observações comerciais para o faturamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   96
            Left            =   9847
            TabIndex        =   342
            Top             =   240
            Width           =   3180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de cobrança"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   84
            Left            =   10612
            TabIndex        =   272
            Top             =   3810
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Validade da proposta comercial"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   83
            Left            =   10147
            TabIndex        =   271
            Top             =   2910
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Garantia do(s) produto(s)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   82
            Left            =   10335
            TabIndex        =   270
            Top             =   2010
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Informações do transporte"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   81
            Left            =   10290
            TabIndex        =   269
            Top             =   1110
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de entrega"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   3120
            TabIndex        =   221
            Top             =   3810
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo para reajuste"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   3015
            TabIndex        =   220
            Top             =   2910
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Impostos inclusos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   3097
            TabIndex        =   219
            Top             =   2010
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Desenvolvimentos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   3075
            TabIndex        =   217
            Top             =   1110
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Condições de pagamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   2820
            TabIndex        =   172
            Top             =   210
            Width           =   1815
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74940
         TabIndex        =   174
         Top             =   360
         Width           =   15255
         _ExtentX        =   26908
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
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   12300
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_PI.frx":1B45F5
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   60
         TabIndex        =   175
         Top             =   330
         Width           =   15270
         _ExtentX        =   26935
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
         ButtonCaption7  =   "Estrutura"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Estrutura (F7)"
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
         ButtonWidth7    =   53
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Emitir PI"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Emitir pedido interno (F8)"
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
         ButtonLeft8     =   323
         ButtonTop8      =   2
         ButtonWidth8    =   47
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Cancelar PI"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Cancelar pedido interno (F9)"
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
         ButtonLeft9     =   372
         ButtonTop9      =   2
         ButtonWidth9    =   63
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Composição"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Composição (F10)"
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
         ButtonLeft10    =   437
         ButtonTop10     =   2
         ButtonWidth10   =   65
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Status"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Alterar status [F11]"
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
         ButtonLeft11    =   504
         ButtonTop11     =   2
         ButtonWidth11   =   39
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Alterações"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Cadastrar alterações."
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
         ButtonLeft12    =   545
         ButtonTop12     =   2
         ButtonWidth12   =   59
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Empenho"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Verificar empenho do item no estoque"
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
         ButtonLeft13    =   606
         ButtonTop13     =   2
         ButtonWidth13   =   52
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
         ButtonLeft14    =   660
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   664
         ButtonTop15     =   2
         ButtonWidth15   =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft16    =   702
         ButtonTop16     =   2
         ButtonWidth16   =   26
         ButtonHeight16  =   21
         ButtonUseMaskColor16=   0   'False
         ButtonEnabled17 =   0   'False
         ButtonIconSize17=   32
         ButtonKey17     =   "16"
         ButtonAlignment17=   2
         BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState17   =   5
         ButtonLeft17    =   730
         ButtonTop17     =   2
         ButtonWidth17   =   24
         ButtonHeight17  =   24
         ButtonUseMaskColor17=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   12660
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_PI.frx":1B8C8F
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74940
         TabIndex        =   177
         Top             =   360
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1720
         ButtonCount     =   15
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
         ButtonCaption7  =   "Emitir PI"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Emitir pedido interno (F7)"
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
         ButtonWidth7    =   47
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Cancelar PI"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Cancelar pedido interno (F8)"
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
         ButtonLeft8     =   317
         ButtonTop8      =   2
         ButtonWidth8    =   63
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Composição"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Composição (F9)"
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
         ButtonLeft9     =   382
         ButtonTop9      =   2
         ButtonWidth9    =   65
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Status"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Alterar status do serviço [F11]"
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
         ButtonLeft10    =   449
         ButtonTop10     =   2
         ButtonWidth10   =   39
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Alterações"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Cadastrar alterações."
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
         ButtonLeft11    =   490
         ButtonTop11     =   2
         ButtonWidth11   =   59
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonAlignment12=   2
         ButtonType12    =   1
         ButtonStyle12   =   -1
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   -1
         ButtonLeft12    =   551
         ButtonTop12     =   4
         ButtonWidth12   =   2
         ButtonHeight12  =   54
         ButtonCaption13 =   "Ajuda"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Ajuda (F1)"
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
         ButtonLeft13    =   555
         ButtonTop13     =   2
         ButtonWidth13   =   36
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Sair"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Sair (Esc)"
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
         ButtonLeft14    =   593
         ButtonTop14     =   2
         ButtonWidth14   =   26
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonKey15     =   "15"
         ButtonAlignment15=   2
         BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState15   =   5
         ButtonLeft15    =   621
         ButtonTop15     =   2
         ButtonWidth15   =   24
         ButtonHeight15  =   24
         ButtonUseMaskColor15=   0   'False
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   12630
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_PI.frx":1C1C19
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar5 
         Height          =   975
         Left            =   -74940
         TabIndex        =   178
         Top             =   360
         Width           =   15285
         _ExtentX        =   26961
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
         ButtonLeft2     =   37
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   75
         ButtonTop3      =   2
         ButtonWidth3    =   38
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
         ButtonLeft4     =   115
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
         ButtonLeft5     =   168
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
         ButtonLeft6     =   217
         ButtonTop6      =   2
         ButtonWidth6    =   46
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
         ButtonLeft7     =   265
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   269
         ButtonTop8      =   2
         ButtonWidth8    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   307
         ButtonTop9      =   2
         ButtonWidth9    =   26
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
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
         ButtonState10   =   5
         ButtonLeft10    =   335
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
         Begin DrawSuite2022.USImageList USImageList5 
            Left            =   13830
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_PI.frx":1CA0C7
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2985
         Left            =   60
         TabIndex        =   224
         Top             =   1350
         Width           =   15315
         _ExtentX        =   27014
         _ExtentY        =   5265
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
         TabCaption(0)   =   "Dados principais"
         TabPicture(0)   =   "frmVendas_PI.frx":1CF336
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1(10)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Chk_CFOP_prod"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Dados adicionais"
         TabPicture(1)   =   "frmVendas_PI.frx":1CF352
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Chk_obs_faturamento_prod"
         Tab(1).Control(1)=   "Frame1(12)"
         Tab(1).ControlCount=   2
         Begin VB.CheckBox Chk_obs_faturamento_prod 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   -65460
            TabIndex        =   238
            Top             =   1080
            Width           =   195
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
            Height          =   2625
            Index           =   12
            Left            =   -74970
            TabIndex        =   237
            Top             =   330
            Width           =   15255
            Begin VB.CheckBox chkNovo_projeto 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Novo projeto?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   13680
               TabIndex        =   410
               Top             =   180
               Width           =   1335
            End
            Begin VB.TextBox txtDureza 
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
               Left            =   7290
               MaxLength       =   50
               TabIndex        =   408
               ToolTipText     =   "Duzera."
               Top             =   390
               Width           =   1815
            End
            Begin VB.CheckBox chkRetorno 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Dt Retorno"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   13770
               TabIndex        =   376
               Top             =   750
               Width           =   1155
            End
            Begin VB.CheckBox Chk_utiliza_mat_consignado 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Utiliza material consignado?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9150
               TabIndex        =   93
               Top             =   420
               Width           =   2445
            End
            Begin VB.TextBox txtGravacao 
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
               Height          =   795
               Left            =   10170
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   98
               TabStop         =   0   'False
               ToolTipText     =   "Gravação."
               Top             =   1590
               Width           =   4485
            End
            Begin VB.TextBox txtembalagem 
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
               Height          =   795
               Left            =   5190
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   97
               TabStop         =   0   'False
               ToolTipText     =   "Embalagem."
               Top             =   1590
               Width           =   4485
            End
            Begin VB.TextBox txtinspecao 
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
               Height          =   795
               Left            =   180
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   96
               TabStop         =   0   'False
               ToolTipText     =   "Inspeção."
               Top             =   1590
               Width           =   4515
            End
            Begin VB.TextBox txtEspessura 
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
               Left            =   1770
               MaxLength       =   30
               TabIndex        =   88
               ToolTipText     =   "Espessura (mm)."
               Top             =   390
               Width           =   1815
            End
            Begin VB.TextBox txtLargura 
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
               Left            =   3600
               MaxLength       =   30
               TabIndex        =   89
               ToolTipText     =   "Largura (mm)."
               Top             =   390
               Width           =   1815
            End
            Begin VB.TextBox txtComprimento 
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
               Left            =   5430
               MaxLength       =   30
               TabIndex        =   90
               ToolTipText     =   "Comprimento (mm)."
               Top             =   390
               Width           =   1815
            End
            Begin VB.TextBox Txt_observacoes_fat_prod 
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
               Height          =   375
               Left            =   7710
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   95
               ToolTipText     =   "Observações de faturamento."
               Top             =   960
               Width           =   5925
            End
            Begin VB.CheckBox Chk_antecipacao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Antecipação de faturamento?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9150
               TabIndex        =   91
               Top             =   180
               Width           =   2445
            End
            Begin VB.CheckBox Chk_faturamento_parcial 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Faturamento parcial?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   11730
               TabIndex        =   92
               Top             =   180
               Width           =   1815
            End
            Begin VB.TextBox Txt_observacoes_prod 
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
               Height          =   375
               Left            =   180
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   94
               ToolTipText     =   "Observações."
               Top             =   960
               Width           =   7400
            End
            Begin VB.TextBox Txt_n_serie 
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
               TabIndex        =   87
               ToolTipText     =   "Número de série."
               Top             =   390
               Width           =   1575
            End
            Begin DrawSuite2022.USButton cmdBotao 
               Height          =   795
               Index           =   0
               Left            =   4710
               TabIndex        =   367
               Top             =   1590
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   1402
               DibPicture      =   "frmVendas_PI.frx":1CF36E
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton cmdBotao 
               Height          =   795
               Index           =   1
               Left            =   9690
               TabIndex        =   368
               Top             =   1590
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   1402
               DibPicture      =   "frmVendas_PI.frx":1CF46A
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton cmdBotao 
               Height          =   795
               Index           =   2
               Left            =   14670
               TabIndex        =   369
               Top             =   1590
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   1402
               DibPicture      =   "frmVendas_PI.frx":1CF566
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin MSMask.MaskEdBox Txt_data_retorno 
               Height          =   315
               Left            =   13740
               TabIndex        =   377
               ToolTipText     =   "Data prevista do retorno."
               Top             =   975
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Enabled         =   0   'False
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Dureza"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   110
               Left            =   7935
               TabIndex        =   409
               Top             =   180
               Width           =   510
            End
            Begin VB.Image Img_calendario_retorno 
               Enabled         =   0   'False
               Height          =   360
               Left            =   14685
               Picture         =   "frmVendas_PI.frx":1CF662
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   945
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Gravação"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   113
               Left            =   12060
               TabIndex        =   297
               Top             =   1380
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Embalagem"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   112
               Left            =   7020
               TabIndex        =   296
               Top             =   1380
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Observações para faturamento"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   111
               Left            =   9765
               TabIndex        =   295
               Top             =   750
               Width           =   2265
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Comprimento / mm"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   109
               Left            =   5670
               TabIndex        =   294
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Largura / mm"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   108
               Left            =   4035
               TabIndex        =   293
               Top             =   180
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Espessura / mm"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   107
               Left            =   2115
               TabIndex        =   292
               Top             =   180
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Inspeção"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   30
               Left            =   2100
               TabIndex        =   242
               Top             =   1380
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "N. de série"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   570
               TabIndex        =   241
               Top             =   180
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Observações"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   18
               Left            =   3408
               TabIndex        =   239
               Top             =   750
               Width           =   945
            End
         End
         Begin VB.CheckBox Chk_CFOP_prod 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8490
            TabIndex        =   236
            Top             =   480
            Width           =   195
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
            Height          =   2625
            Index           =   10
            Left            =   30
            TabIndex        =   225
            Top             =   330
            Width           =   15225
            Begin VB.TextBox txtvFrete 
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
               ForeColor       =   &H00000080&
               Height          =   315
               Left            =   2910
               MaxLength       =   15
               TabIndex        =   374
               TabStop         =   0   'False
               ToolTipText     =   "Valor do frete"
               Top             =   2205
               Width           =   1230
            End
            Begin VB.ComboBox Cmb_prioridade 
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
               ItemData        =   "frmVendas_PI.frx":1CFAE5
               Left            =   5220
               List            =   "frmVendas_PI.frx":1CFAEF
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   65
               ToolTipText     =   "Prioridade."
               Top             =   1560
               Width           =   1050
            End
            Begin VB.CheckBox Chk_prazo_prod 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   13710
               TabIndex        =   229
               Top             =   1350
               Width           =   195
            End
            Begin VB.TextBox txtpccliente 
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
               Left            =   8850
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   69
               TabStop         =   0   'False
               ToolTipText     =   "Pedido do cliente."
               Top             =   1560
               Width           =   1335
            End
            Begin VB.CheckBox Chk_PC_prod 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   8910
               TabIndex        =   228
               Top             =   1350
               Visible         =   0   'False
               Width           =   195
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
               Left            =   6270
               Locked          =   -1  'True
               TabIndex        =   66
               TabStop         =   0   'False
               ToolTipText     =   "ID da NCM."
               Top             =   1560
               Width           =   525
            End
            Begin VB.TextBox Txt_n_item 
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
               Left            =   10200
               MaxLength       =   6
               TabIndex        =   71
               ToolTipText     =   "Número do produto/item no pedido do cliente."
               Top             =   1560
               Width           =   675
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
               ItemData        =   "frmVendas_PI.frx":1CFB04
               Left            =   8100
               List            =   "frmVendas_PI.frx":1CFBBF
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   68
               ToolTipText     =   "Situação tributária ICMS."
               Top             =   1560
               Width           =   750
            End
            Begin VB.ComboBox Cmb_un_com 
               Appearance      =   0  'Flat
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
               Left            =   855
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   63
               ToolTipText     =   "Unidade comercial."
               Top             =   1560
               Width           =   675
            End
            Begin VB.TextBox Txt_ID_CFOP_prod 
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
               Left            =   8400
               Locked          =   -1  'True
               TabIndex        =   55
               TabStop         =   0   'False
               ToolTipText     =   "ID da CFOP."
               Top             =   375
               Width           =   525
            End
            Begin VB.TextBox Txt_natureza_operacao_prod 
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
               Left            =   9630
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   57
               TabStop         =   0   'False
               ToolTipText     =   "Descrição da natureza da operação."
               Top             =   375
               Width           =   5085
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
               Height          =   315
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   52
               ToolTipText     =   "Código interno."
               Top             =   375
               Width           =   1320
            End
            Begin VB.TextBox txtvalor_total 
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
               Left            =   13170
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   85
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   2205
               Width           =   1905
            End
            Begin VB.TextBox txtdbl_valoripi 
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
               Left            =   10365
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   82
               TabStop         =   0   'False
               ToolTipText     =   "Valor do IPI."
               Top             =   2205
               Width           =   1065
            End
            Begin VB.TextBox txtvalorunitariodesc 
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
               Left            =   8025
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   80
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   2205
               Width           =   1335
            End
            Begin VB.TextBox txtvalordesconto 
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
               Left            =   6675
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   79
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   2205
               Width           =   1335
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
               Height          =   315
               Left            =   5640
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   77
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   2205
               Width           =   1020
            End
            Begin VB.ComboBox cmbfamilia 
               Appearance      =   0  'Flat
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
               Left            =   1545
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   64
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   1560
               Width           =   3660
            End
            Begin VB.ComboBox cmbun 
               Appearance      =   0  'Flat
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
               Left            =   180
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   62
               ToolTipText     =   "Unidade de estoque."
               Top             =   1560
               Width           =   675
            End
            Begin VB.TextBox txtvalorunitario 
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
               ForeColor       =   &H00000080&
               Height          =   315
               Left            =   1500
               MaxLength       =   15
               TabIndex        =   75
               ToolTipText     =   "Valor unitário."
               Top             =   2205
               Width           =   1395
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
               ForeColor       =   &H00000080&
               Height          =   315
               Left            =   180
               MaxLength       =   15
               TabIndex        =   74
               ToolTipText     =   "Quantidade."
               Top             =   2205
               Width           =   975
            End
            Begin VB.TextBox txtint_icms 
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
               Left            =   11445
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   83
               TabStop         =   0   'False
               ToolTipText     =   "Porcentagem do ICMS."
               Top             =   2205
               Width           =   465
            End
            Begin VB.TextBox txtInt_ipi 
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
               Left            =   9885
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   81
               TabStop         =   0   'False
               ToolTipText     =   "Porcentagem do IPI."
               Top             =   2205
               Width           =   465
            End
            Begin VB.TextBox txtvalor_icms 
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
               Left            =   11925
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   84
               TabStop         =   0   'False
               ToolTipText     =   "Valor do ICMS."
               Top             =   2205
               Width           =   1230
            End
            Begin VB.TextBox txtEspecificacoes 
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
               Left            =   8580
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               ToolTipText     =   "Descrição comercial."
               Top             =   960
               Width           =   6465
            End
            Begin VB.TextBox txtdesctecnica 
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
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   105
               TabIndex        =   60
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   960
               Width           =   8385
            End
            Begin VB.TextBox txtRev_cod 
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
               Left            =   3735
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   53
               TabStop         =   0   'False
               Text            =   "0"
               ToolTipText     =   "Revisão."
               Top             =   375
               Width           =   525
            End
            Begin VB.ComboBox cmbreferencia 
               Appearance      =   0  'Flat
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
               Left            =   5340
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   54
               ToolTipText     =   "Código de referência."
               Top             =   375
               Width           =   1890
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto código?"
               ForeColor       =   &H00000080&
               Height          =   555
               Index           =   5
               Left            =   150
               TabIndex        =   227
               Top             =   150
               Width           =   2235
               Begin VB.CheckBox OPTnovoman 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Manual"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   1350
                  TabIndex        =   59
                  Top             =   240
                  Width           =   825
               End
               Begin VB.CheckBox OPTnovo 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Automático "
                  ForeColor       =   &H00800000&
                  Height          =   225
                  Left            =   180
                  TabIndex        =   58
                  Top             =   240
                  Width           =   1305
               End
            End
            Begin VB.TextBox Txt_analise 
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
               Left            =   11880
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   70
               TabStop         =   0   'False
               ToolTipText     =   "Análise crítica."
               Top             =   1560
               Width           =   1305
            End
            Begin VB.TextBox txtComissao 
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
               ForeColor       =   &H00000080&
               Height          =   315
               Left            =   4150
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   86
               ToolTipText     =   "Comissão do vendedor."
               Top             =   2205
               Width           =   915
            End
            Begin VB.CheckBox Chk_desc 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5640
               TabIndex        =   76
               Top             =   1995
               Width           =   225
            End
            Begin VB.CheckBox Chk_valor_desc 
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6690
               TabIndex        =   78
               Top             =   1995
               Width           =   225
            End
            Begin VB.TextBox Txt_CFOP_prod 
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
               Left            =   8940
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   56
               TabStop         =   0   'False
               ToolTipText     =   "Natureza da operação."
               Top             =   375
               Width           =   675
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
               Left            =   6810
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   67
               TabStop         =   0   'False
               ToolTipText     =   "Classificação fiscal."
               Top             =   1560
               Width           =   945
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
               Height          =   315
               Left            =   5340
               MaxLength       =   50
               TabIndex        =   226
               ToolTipText     =   "Código de referência."
               Top             =   375
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.TextBox txtPrazo_Produto 
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
               Height          =   315
               Left            =   13680
               MaxLength       =   4
               TabIndex        =   72
               ToolTipText     =   "Prazo de entrega em dias."
               Top             =   1560
               Width           =   1095
            End
            Begin MSMask.MaskEdBox mskprazo 
               Height          =   315
               Left            =   13680
               TabIndex        =   73
               ToolTipText     =   "Prazo de entrega."
               Top             =   1560
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin DrawSuite2022.USButton cmdlistaproduto 
               Height          =   315
               Left            =   4620
               TabIndex        =   357
               Top             =   390
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1CFD0F
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
               BorderColorDisabled=   12632256
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               GradientColor1  =   12632256
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
            Begin DrawSuite2022.USButton Cmd_visualizar_arquivo 
               Height          =   315
               Left            =   4950
               TabIndex        =   358
               Top             =   390
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1CFE0B
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton cmdfiltrar 
               Height          =   315
               Left            =   4290
               TabIndex        =   359
               Top             =   390
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D345B
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
               GradientColor1  =   12632256
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
            Begin DrawSuite2022.USButton Cmd_localizar_CFOP_prod 
               Height          =   315
               Left            =   14730
               TabIndex        =   360
               Top             =   390
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D3726
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton cmdCF 
               Height          =   315
               Left            =   7770
               TabIndex        =   361
               Top             =   1560
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D3822
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton Cmd_importar_PC_prod 
               Height          =   315
               Left            =   10890
               TabIndex        =   362
               Top             =   1560
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D391E
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton Cmd_limpar_caminho_PC_Prod 
               Height          =   315
               Left            =   11220
               TabIndex        =   363
               Top             =   1560
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D3A1A
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton Cmb_visualizar_PC_prod 
               Height          =   315
               Left            =   11550
               TabIndex        =   364
               Top             =   1560
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D63BB
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton Cmd_analise 
               Height          =   315
               Left            =   13200
               TabIndex        =   365
               Top             =   1560
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D6686
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
               GradientColor1  =   12632256
               GradientColor2  =   12632256
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
            Begin DrawSuite2022.USButton cmdcalc_peso 
               Height          =   315
               Left            =   1170
               TabIndex        =   366
               Top             =   2210
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmVendas_PI.frx":1D67CB
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
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   10890
               Top             =   270
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   141
               Left            =   12150
               TabIndex        =   373
               Top             =   1995
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   140
               Left            =   11490
               TabIndex        =   372
               Top             =   1995
               Width           =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Frete"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   9
               Left            =   3150
               TabIndex        =   371
               Top             =   1995
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "n° item"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   104
               Left            =   10290
               TabIndex        =   370
               Top             =   1350
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição comercial"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   118
               Left            =   11085
               TabIndex        =   302
               Top             =   750
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "% Comissão"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   106
               Left            =   4140
               TabIndex        =   291
               Top             =   1995
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total produtos"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   105
               Left            =   13440
               TabIndex        =   290
               Top             =   1995
               Width           =   1425
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor IPI"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   102
               Left            =   10605
               TabIndex        =   289
               Top             =   1995
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "IPI"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   101
               Left            =   10005
               TabIndex        =   288
               Top             =   1995
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor unit. c/ desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   100
               Left            =   8025
               TabIndex        =   287
               Top             =   1995
               Width           =   1350
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor do desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   99
               Left            =   6915
               TabIndex        =   286
               Top             =   1995
               Width           =   1020
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Desc. (%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   98
               Left            =   5880
               TabIndex        =   285
               Top             =   1995
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor unitário*"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   97
               Left            =   1695
               TabIndex        =   284
               Top             =   1995
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "N° Análise crítica"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   95
               Left            =   11940
               TabIndex        =   283
               Top             =   1350
               Width           =   1200
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido cliente"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   94
               Left            =   9120
               TabIndex        =   282
               Top             =   1350
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "CST ICMS"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   93
               Left            =   8130
               TabIndex        =   281
               Top             =   1350
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "NCM*"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   92
               Left            =   7110
               TabIndex        =   280
               Top             =   1350
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   91
               Left            =   6450
               TabIndex        =   279
               Top             =   1350
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Família"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   90
               Left            =   3135
               TabIndex        =   278
               Top             =   1350
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   88
               Left            =   915
               TabIndex        =   277
               Top             =   1350
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Natureza da operação*"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   89
               Left            =   11370
               TabIndex        =   276
               Top             =   180
               Width           =   1695
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código de referência"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   87
               Left            =   5475
               TabIndex        =   275
               Top             =   180
               Width           =   1500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   86
               Left            =   8670
               TabIndex        =   274
               Top             =   180
               Width           =   165
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CFOP*"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   85
               Left            =   9090
               TabIndex        =   273
               Top             =   180
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Prioridade"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   5355
               TabIndex        =   240
               Top             =   1350
               Width           =   720
            End
            Begin VB.Image imgCalendario 
               Height          =   360
               Left            =   14760
               Picture         =   "frmVendas_PI.frx":1D902E
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   1530
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo final"
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
               Index           =   27
               Left            =   13965
               TabIndex        =   235
               Top             =   1350
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Index           =   25
               Left            =   2460
               TabIndex        =   234
               Top             =   180
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   19
               Left            =   240
               TabIndex        =   233
               Top             =   1350
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantidade*"
               ForeColor       =   &H00000080&
               Height          =   195
               Index           =   20
               Left            =   240
               TabIndex        =   232
               Top             =   1995
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rev."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   16
               Left            =   3840
               TabIndex        =   231
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição técnica"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   17
               Left            =   4020
               TabIndex        =   230
               Top             =   750
               Width           =   1245
            End
         End
      End
      Begin DrawSuite2022.USProgressBar USProgressBar1 
         Height          =   255
         Left            =   60
         TabIndex        =   375
         Top             =   9690
         Width           =   15240
         _ExtentX        =   26882
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
   End
End
Attribute VB_Name = "frmVendas_PI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_PI                  As Boolean 'OK
Public Novo_PI1                 As Boolean 'OK
Public Novo_PI2                 As Boolean 'OK
Dim Novo_PI3                    As Boolean 'OK
Public Produto_servico          As Boolean 'OK
Public IDAnalise                As Integer 'OK
Public IDAnalise_servico        As Integer 'OK
Public StrSql_PI_Localizar      As String 'OK
Dim TBLISTA_Vendas_PI           As ADODB.Recordset 'OK
Dim TBLISTA_Vendas_PI1          As ADODB.Recordset 'OK
Dim TBLISTA_Vendas_PI2          As ADODB.Recordset 'OK
Dim Caminho_PC_prod_PI          As String
Dim Caminho_PC_serv_PI          As String
Public StrSql_PI_LocalizarRel   As String 'OK
Public TabelaSN_PI As Integer 'OK
Public RegimeEmpresa_PI As Integer 'OK

Private Sub ProcLocalizaVendedorInterno()
On Error GoTo tratar_erro

If txtIDcliente = "" Then Exit Sub

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from empresa where Empresa = '" & Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAcessos.EOF = False Then
ClienteVendedor = TBAcessos!ClienteVendedor
SemEstoque = TBAcessos!ClienteVendedor
End If
TBAcessos.Close

'===============================================================
' Se utilizar clientes por vendedor
'===============================================================
If ClienteVendedor = True Then
Inicio:
'===============================================================
' Verifica se esse cliente pertence ao vendedor
'===============================================================
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Vendas_Vendedores_Clientes where IDCliente = " & txtIDcliente & "", Conexao, adOpenKeyset, adLockOptimistic
'===============================================================
' Se pertencer ao vendedor
'===============================================================
If TBClientes.EOF = False Then

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * FROM Vendas_Vendedores WHERE ID = " & TBClientes!IDvendedor & "", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = True Then
USMsgBox "Atenção esse cliente não tem vendedor interno, faça seu cadastro"
Exit Sub
End If

txtComissao.Text = IIf(IsNull(TBUsuarios!Comissao), 0, TBUsuarios!Comissao)
txtVI.Text = TBUsuarios!ID
txtvend_Int.Text = TBUsuarios!vendedor
TBClientes.Close
TBUsuarios.Close
End If
Else
'================================================================
' Se não controlar cliente por vendedor permite escolher o vendedor interno
'================================================================
cmdVendedor_Interno_Click
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAjuda()
On Error GoTo tratar_erro

If Vendas_PI = True Then
    FunAbrirVideoWeb ("http://www.youtube.com/watch?v=KxGZkioTqlg&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=24&feature=plcp")
Else
    FunAbrirVideoWeb ("http://www.youtube.com/watch?v=Y17NVW6Fenc&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=32&feature=plcp")
End If

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

Private Sub BtnRedespacho_Click()
On Error GoTo tratar_erro

Transporte1 = False
Transporte2 = True
frmVendas_Transporte_Tipo.Show 1
Vendas_Proposta = False
Vendas_PI = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSalvarFrete_Click()
On Error GoTo tratar_erro

If USMsgBox("Atenção!" & vbCrLf & "Deseja realmente modificar o valor do frete dos itens da lista?", vbYesNo, "CAPRIND v5.0") = vbYes Then

If TxtTotalFrete.Text <> "" Then

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_carteira where cotacao = " & txtId.Text & " and TIPO = 'P'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast
Contador = TBLISTA.RecordCount
TBLISTA.MoveFirst
vFrete = TxtTotalFrete / Contador
Do While TBLISTA.EOF = False
TBLISTA!vFrete = TxtTotalFrete / Contador
TBLISTA!dbl_Valor_ICMS = ((TBLISTA!preco_lote + vFrete) * TBLISTA!IntICMS) / 100
TBLISTA!BC_ICMS = (TBLISTA!preco_unitario * TBLISTA!quantidade) + vFrete
TBLISTA.Update
TBLISTA.MoveNext
Loop
End If
TBLISTA.Close
End If
ProcAtualizalistaProdutos 1

ProcGravarTotais IIf(txtId = "", 0, txtId)
ProcPuxaTotais

ProcLimparProdutos True
USMsgBox ("Valores dos itens atualizados com sucesso!"), vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnTransportadora_Click()
On Error GoTo tratar_erro

Transporte1 = True
Transporte2 = False
Vendas_Proposta = False
Vendas_PI = True
frmVendas_Transporte_Tipo.Show 1

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
    .Text = ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_desc2_Click()
On Error GoTo tratar_erro

With txtdesconto2
    If Chk_desc2.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_valor_desc2.Value = 0
        txtvalordesconto2.Locked = True
        txtvalordesconto2.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
    .Text = ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_servico_executado_cliente_Click()
On Error GoTo tratar_erro

ProcCarregaISSQN

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
    .Text = ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_valor_desc2_Click()
On Error GoTo tratar_erro

With txtvalordesconto2
    If Chk_valor_desc2.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_desc2.Value = 0
        txtdesconto2.Locked = True
        txtdesconto2.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
    .Text = ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkRetorno_Click()
On Error GoTo tratar_erro

If chkRetorno.Value = 1 Then
    Txt_data_retorno.Enabled = True
    Img_calendario_retorno.Enabled = True
Else
    With Txt_data_retorno
        .Text = "__/__/____"
        .Enabled = False
    End With
    Img_calendario_retorno.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_CST_ICMS_Click()
On Error GoTo tratar_erro

ProcAtualizavalores

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub Cmb_empresa_Change()
On Error GoTo tratar_erro

IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    Select Case Cmb_opcao_lista
        Case "Validação"
            .ButtonState(13) = 5
            .ButtonState(14) = 0
        Case "Status"
            .ButtonState(13) = 0
            .ButtonState(14) = 5
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmb_Tipo_Frete_Change()
On Error GoTo tratar_erro

If cmb_Tipo_Frete.Text = "EMITENTE (CIF)" Then
FRETE_ICMS = True
Else
FRETE_ICMS = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmb_Tipo_Frete_Click()
On Error GoTo tratar_erro

If cmb_Tipo_Frete.Text = "EMITENTE (CIF)" Then
FRETE_ICMS = True
Else
FRETE_ICMS = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmb_visualizar_PC_prod_Click()
On Error GoTo tratar_erro

If Caminho_PC_prod_PI <> "" Then ProcAbrirArquivo Caminho_PC_prod_PI

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_visualizar_PC_serv_Click()
On Error GoTo tratar_erro

If Caminho_PC_serv_PI <> "" Then ProcAbrirArquivo Caminho_PC_serv_PI

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

txtdesctecnica = FunBuscaDescPadraoFamilia(cmbfamilia, txtNomenclatura, txtdesctecnica)
Txt_ID_CF = FunBuscaIDCFPadraoFamilia(cmbfamilia, txtNomenclatura, IIf(Txt_ID_CF = "", 0, Txt_ID_CF))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfamiliaservico_Click()
On Error GoTo tratar_erro

txtdescservico = FunBuscaDescPadraoFamilia(cmbfamiliaservico, txtcodservico, txtdescservico)

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

Private Sub cmbOpcao_lista_prod_Click()
On Error GoTo tratar_erro

With Listprod
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar3
    Select Case cmbOpcao_lista_prod
        Case "Excluir"
            .ButtonState(3) = 0
            .ButtonState(11) = 5
        Case "Status"
            .ButtonState(3) = 5
            .ButtonState(11) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbOpcao_lista_serv_Click()
On Error GoTo tratar_erro

With ListaServicos
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar4
    Select Case cmbOpcao_lista_serv
        Case "Excluir"
            .ButtonState(3) = 0
            .ButtonState(10) = 5
        Case "Status"
            .ButtonState(3) = 5
            .ButtonState(10) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbun_Click()
On Error GoTo tratar_erro

If cmbun <> "" Then ProcLibera_UN_Com cmbun, Cmb_un_com

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_analise_Click()
On Error GoTo tratar_erro

Vendas_Produtos = True
frmVendas_propostaII_ListaAnalise.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_analise1_Click()
On Error GoTo tratar_erro

Vendas_Produtos = False
frmVendas_propostaII_ListaAnalise.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEstrutura()
On Error GoTo tratar_erro

If txtNomenclatura = "" Then
    USMsgBox ("Informe o produto antes de abrir a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmproj_conjunto
    .Show
    .Lista.ListItems.Clear
    .ProcLimpaCampos
    .ProcLimpaCamposItem
    .Procatualizadados (txtNomenclatura)
    .ProcCarregaVersao ""
    .Novo_Conjunto = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImpostos()
On Error GoTo tratar_erro

If txtCotacao = "" Then
    USMsgBox ("Informe " & IIf(Vendas_Proposta = True, "a proposta", "o pedido") & " antes de visualizar os impostos."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcVerificaEmpresaCliente
FrmImpostos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_importar_PC_serv_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Caminho_PC_serv_PI = caminho
If Vendas_PI = True And Caminho_PC_serv_PI <> "" Then If USMsgBox("Deseja replicar o nome do arquivo para o número do pedido de compra?", vbYesNo, "CAPRIND v5.0") = vbYes Then txtpcclienteserv = Nome_anexo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_limpar_caminho_PC_prod_Click()
On Error GoTo tratar_erro

Caminho_PC_prod_PI = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_importar_PC_prod_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Caminho_PC_prod_PI = caminho
If Vendas_PI = True And Caminho_PC_prod_PI <> "" Then If USMsgBox("Deseja replicar o nome do arquivo para o número do pedido de compra?", vbYesNo, "CAPRIND v5.0") = vbYes Then txtpccliente = Nome_anexo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_limpar_caminho_PC_serv_Click()
On Error GoTo tratar_erro

Caminho_PC_serv_PI = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CF_Click()
On Error GoTo tratar_erro

Txt_ID_CF = ""
Txt_CF = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CFOP_Click()
On Error GoTo tratar_erro

If txtDtValidacao <> "" Then Exit Sub
Txt_ID_CFOP_prod = ""
Txt_CFOP_prod = ""
Txt_natureza_operacao_prod = ""
chkRetorno.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CFOP_serv_Click()
On Error GoTo tratar_erro

Txt_ID_CFOP_serv = ""
Txt_CFOP_serv = ""
Txt_natureza_operacao_serv = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_CFOP_prod_Click()
On Error GoTo tratar_erro

If txtDtValidacao <> "" Then Exit Sub
Clientes = False
'Vendas_Proposta = True
'Vendas_PI = False
Faturamento = False
Compras_Pedido = False
Sit_REG = 1
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_CFOP_serv_Click()
On Error GoTo tratar_erro

Clientes = False
'Vendas_Proposta = True
'Vendas_PI = False
Faturamento = False
Compras_Pedido = False
Sit_REG = 2
frm_ListaNatureza.Show 1

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

If txtcodservico = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtcodservico & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdBotao_Click(index As Integer)
On Error GoTo tratar_erro

Select Case index
    Case 0: Aplic = 4
    Case 1: Aplic = 5
    Case 2: Aplic = 12
End Select

Compras_Cotacao = False
Compras_Pedido = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdBotao_DadosComerciais_Click(index As Integer)
On Error GoTo tratar_erro

Select Case index
    Case 0: Aplic = 1
    Case 1: Aplic = 3
    Case 2: Aplic = 6
    Case 3: Aplic = 7
    Case 4: Aplic = 8
    Case 5: Aplic = 9
    Case 6: Aplic = 10
End Select

Compras_Cotacao = False
Compras_Pedido = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcalc_peso_Click()
On Error GoTo tratar_erro

If txtNomenclatura = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = False
    Compras_Cotacao = False
    Compras_Pedido = False
    Estoque_recebimento = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

'Verifica se o produto pertence ao cliente
IDCliente = txtIDcliente
Cliente = txtCliente
Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select * from projproduto where desenho = '" & txtNomenclatura.Text & "' and Vendas = 'True' and Tipo = 'P' and Bloqueado = 'False' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & TBCompras_Pedido!Codproduto & " and IDCliente <> 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Projproduto_clientes where Codproduto = " & TBCompras_Pedido!Codproduto & " and IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from empresa where codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloquear_produtos = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                USMsgBox ("Este produto não pertence ao cliente " & Cliente & "."), vbExclamation, "CAPRIND v5.0"
                ProcLimparProdutos True
                TBCiclo.Close
                Exit Sub
            Else
                If USMsgBox("Este produto não pertence ao cliente " & Cliente & ", deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                    ProcLimparProdutos True
                    TBCiclo.Close
                    Exit Sub
                End If
            End If
            TBCiclo.Close
        End If
        TBProduto.Close
    End If
    TBAbrir.Close
End If
TBCompras_Pedido.Close
ProcPuxaDadosProduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaEmpresaCliente()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where Simples = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Permitido = True Else Permitido = False

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Clientes where IDCliente = " & txtIDcliente & " and Simples = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Permitido = True Else Permitido = False
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosProduto()
On Error GoTo tratar_erro

If txtNomenclatura.Text <> "" Then
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from projproduto where desenho = '" & txtNomenclatura.Text & "' and vendas = 'True' and (tipo = 'P' or tipo = 'PI') and Bloqueado = 'False' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        txtNomenclatura = TBCompras_Pedido!Desenho
        txtEspecificacoes.Text = IIf(IsNull(TBCompras_Pedido!descricaotecnica), "", TBCompras_Pedido!descricaotecnica)
        Txt_observacoes_prod = IIf(IsNull(TBCompras_Pedido!Observacoes), "", TBCompras_Pedido!Observacoes)
        txtespessura = IIf(IsNull(TBCompras_Pedido!Espessura), "", TBCompras_Pedido!Espessura)
        txtLargura = IIf(IsNull(TBCompras_Pedido!Largura), "", TBCompras_Pedido!Largura)
        txtComprimento = IIf(IsNull(TBCompras_Pedido!Comprimento), "", TBCompras_Pedido!Comprimento)
        'txtDureza = IIf(IsNull(TBCompras_Pedido!Dureza), "", TBCompras_Pedido!Dureza)
        txtdesctecnica.Text = IIf(IsNull(TBCompras_Pedido!Descricao), "", TBCompras_Pedido!Descricao)
        txtRev_cod = IIf(IsNull(TBCompras_Pedido!RevDesenho), "", TBCompras_Pedido!RevDesenho)
        txtinspecao = IIf(IsNull(TBCompras_Pedido!Inspecao), "", TBCompras_Pedido!Inspecao)
        txtembalagem = IIf(IsNull(TBCompras_Pedido!Embalagem), "", TBCompras_Pedido!Embalagem)
        txtGravacao = IIf(IsNull(TBCompras_Pedido!Gravacao), "", TBCompras_Pedido!Gravacao)
        
        cmbun.ListIndex = -1
        NomeCampo = "a unidade de estoque"
        If IsNull(TBCompras_Pedido!Unidade) = False And TBCompras_Pedido!Unidade <> "" <> "" Then cmbun = TBCompras_Pedido!Unidade
        Cmb_un_com.ListIndex = -1
        NomeCampo = "a unidade comercial"
        If IsNull(TBCompras_Pedido!Unidade_com) = False And TBCompras_Pedido!Unidade_com <> "" <> "" Then Cmb_un_com = TBCompras_Pedido!Unidade_com
        cmbfamilia.ListIndex = -1
        NomeCampo = "a família"
        If IsNull(TBCompras_Pedido!Classe) = False And TBCompras_Pedido!Classe <> "" Then cmbfamilia = TBCompras_Pedido!Classe
        Txt_ID_CF = ""
        Txt_CF = ""
        
1:
        valor = IIf(Txt_valor_moeda = "", 1, Txt_valor_moeda)
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & TBCompras_Pedido!Codproduto & " and idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If txttipocliente <> "JR" And txttipocliente <> "FR" Then
                txtvalorunitario = IIf(IsNull(TBAbrir!PConsumo), "", Format((TBAbrir!PConsumo / FunVerificaTabelaConversaoUnidade(cmbun, Cmb_un_com)) / valor, "###,##0.0000000000"))
            Else
                txtvalorunitario = IIf(IsNull(TBAbrir!PRevenda), "", Format((TBAbrir!PRevenda / FunVerificaTabelaConversaoUnidade(cmbun, Cmb_un_com)) / valor, "###,##0.0000000000"))
            End If
            If IsNull(TBAbrir!ID_CF) = False Then
                Txt_ID_CF = TBAbrir!ID_CF
            Else
                If IsNull(TBCompras_Pedido!ID_CF) = False Then Txt_ID_CF = TBCompras_Pedido!ID_CF
            End If
        Else
            If txttipocliente <> "JR" And txttipocliente <> "FR" Then
                txtvalorunitario = IIf(IsNull(TBCompras_Pedido!PConsumo), "", Format((TBCompras_Pedido!PConsumo / FunVerificaTabelaConversaoUnidade(cmbun, Cmb_un_com)) / valor, "###,##0.0000000000"))
            Else
                txtvalorunitario = IIf(IsNull(TBCompras_Pedido!PRevenda), "", Format((TBCompras_Pedido!PRevenda / FunVerificaTabelaConversaoUnidade(cmbun, Cmb_un_com)) / valor, "###,##0.0000000000"))
            End If
            If IsNull(TBCompras_Pedido!ID_CF) = False Then Txt_ID_CF = TBCompras_Pedido!ID_CF
        End If
        TBAbrir.Close
        
        ProcCarregaDadosCFOPProdServ IIf(IsNull(TBCompras_Pedido!ID_CFOP1), 0, TBCompras_Pedido!ID_CFOP1), True
        
        If Txt_ID_CF <> "" Then
            ProcValorImposto txtCotacao, Txt_ID_CF, IIf(txtIDcliente = "", 0, txtIDcliente), txtCliente, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, IIf(IsNull(TBCompras_Pedido!ID_CFOP1), 0, TBCompras_Pedido!ID_CFOP1), RegimeEmpresa_PI
            ProcControleImposto IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), IIf(txtIDcliente = "", 0, txtIDcliente)
            If TemIPI = "SIM" Then txtInt_ipi = IntIPI Else txtInt_ipi.Text = 0
            If TemICMS = "SIM" Then txtint_icms = IntICMS Else txtint_icms.Text = 0
            
            If txtuf <> "" Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Carregar_CFOP_ST = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    ProcVerifCFOPST Txt_ID_CF, txtuf
                    If Valido = True Then
                        Txt_ID_CFOP_prod = IDAntigo
                        Txt_CFOP_prod = FamiliaAntiga
                        Txt_natureza_operacao_prod = Familiatext
                        Cmb_CST_ICMS = Letra
                    End If
                End If
            End If
        End If
        
        ProcCarregaComboCodRef cmbReferencia, "P.codproduto = " & TBCompras_Pedido!Codproduto, txtIDcliente, "C", True, True
        
        'Carrega comissão
        Set TBExecucao = CreateObject("adodb.recordset")
        TBExecucao.Open "select * from Vendas_Vendedores where N_Vendedor = " & txtVE, Conexao, adOpenKeyset, adLockOptimistic
        If TBExecucao.EOF = False Then
            If TBExecucao!tipocomissao <> "" And IsNull(TBExecucao!tipocomissao) = False Then
                If TBExecucao!tipocomissao = "V" Then
                    txtComissao = TBExecucao!Comissao
                ElseIf TBExecucao!tipocomissao = "MT" Then
                    txtComissao = TBExecucao!Comissao
                Else
                'Debug.print TBExecucao!tipocomissao
                    Set TBCFOP = CreateObject("adodb.recordset")
                    If TBExecucao!tipocomissao = "C" Then TBCFOP.Open "select * from Vendas_Vendedores_Clientes where IDVendedor = " & TBExecucao!ID & " and IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao!tipocomissao = "P" Then TBCFOP.Open "select * from Vendas_Vendedores_Produto where IDVendedor = " & TBExecucao!ID & " and IDProduto = " & TBCompras_Pedido!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao!tipocomissao = "CP" Then TBCFOP.Open "select Vendas_Vendedores_Produto.comissao from Vendas_Vendedores_Produto INNER JOIN Vendas_Vendedores_Clientes on Vendas_Vendedores_Produto.idcliente = Vendas_Vendedores_Clientes.Id where Vendas_Vendedores_Produto.IDVendedor = " & TBExecucao!ID & " and Vendas_Vendedores_Produto.IDProduto = " & TBCompras_Pedido!Codproduto & " and Vendas_Vendedores_clientes.IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                                        
                    If TBCFOP.EOF = False Then txtComissao = TBCFOP!Comissao
                    TBCFOP.Close
                End If
            End If
        End If
        TBExecucao.Close
        
        ProcBloqueiaTabsProd
        With txtvalorunitario
            If TBCompras_Pedido!Valor_bloqueado = True Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
    Else
        USMsgBox ("Não foi encontrado nenhum produto validado com este código interno ou o mesmo está bloqueado."), vbExclamation, "CAPRIND v5.0"
        ProcLimparProdutos True
        ProcLiberaTabsProd
    End If
    TBCompras_Pedido.Close
    ProcCalculaDesconto
    ProcCalculaValores
Else
    ProcLimparProdutos False
    ProcLiberaTabsProd
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste produto."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaTabsProd()
On Error GoTo tratar_erro

cmdlistaproduto.TabStop = False
cmbReferencia.TabStop = False
OPTnovo.TabStop = False
OPTnovoman.TabStop = False
txtdesctecnica.TabStop = False
txtEspecificacoes.TabStop = False
Txt_observacoes_prod.TabStop = False
Cmb_prioridade.TabStop = False
Txt_observacoes_fat_prod.TabStop = False
txtespessura.TabStop = False
txtLargura.TabStop = False
txtComprimento.TabStop = False
'txtDureza.TabStop = False
cmdCF.TabStop = False
cmbun.TabStop = False
If cmbun <> "KG" And cmbun <> "MM" And cmbun <> "MT" And cmbun <> "PC" And cmbun <> "PÇ" Then Cmb_un_com.TabStop = False
cmbfamilia.TabStop = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaTabsProd()
On Error GoTo tratar_erro

cmdlistaproduto.TabStop = True
cmbReferencia.TabStop = True
OPTnovo.TabStop = True
OPTnovoman.TabStop = True
txtdesctecnica.TabStop = True
txtEspecificacoes.TabStop = True
Txt_observacoes_prod.TabStop = True
Cmb_prioridade.TabStop = True
Txt_observacoes_fat_prod.TabStop = True
txtespessura.TabStop = True
txtLargura.TabStop = True
txtComprimento.TabStop = True
'txtDureza.TabStop = True
cmdCF.TabStop = True
cmbun.TabStop = True
Cmb_un_com.TabStop = True
cmbfamilia.TabStop = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdfiltrar_serv_Click()
On Error GoTo tratar_erro

'Verifica se o serviço pertence ao cliente
IDCliente = txtIDcliente
Cliente = txtCliente
Set TBCompras_Pedido = CreateObject("adodb.recordset")
TBCompras_Pedido.Open "Select * from projproduto where desenho = '" & txtcodservico.Text & "' and Vendas = 'True' and Tipo = 'S' and Bloqueado = 'False' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Pedido.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & TBCompras_Pedido!Codproduto & " and IDCliente <> 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Projproduto_clientes where Codproduto = " & TBCompras_Pedido!Codproduto & " and IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from empresa where codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloquear_produtos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                If USMsgBox("Este serviço não pertence ao cliente " & Cliente & ", deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                    ProcLimparServicos True
                    TBCiclo.Close
                    Exit Sub
                End If
            Else
                USMsgBox ("Este serviço não pertence ao cliente " & Cliente & "."), vbExclamation, "CAPRIND v5.0"
                ProcLimparServicos True
                TBCiclo.Close
                Exit Sub
            End If
            TBCiclo.Close
        End If
        TBProduto.Close
    End If
    TBAbrir.Close
End If
TBCompras_Pedido.Close
ProcPuxadadosServico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadadosServico()
On Error GoTo tratar_erro

If txtcodservico.Text <> "" Then
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from projproduto where desenho = '" & txtcodservico.Text & "' and vendas = 'True' and tipo = 'S' and Bloqueado = 'False' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        txtcodservico = TBCompras_Pedido!Desenho
        txtdescservico.Text = IIf(IsNull(TBCompras_Pedido!Descricao), "", TBCompras_Pedido!Descricao)
        txtdesccomservico.Text = IIf(IsNull(TBCompras_Pedido!descricaotecnica), "", TBCompras_Pedido!descricaotecnica)
        txtObs_serv = IIf(IsNull(TBCompras_Pedido!Observacoes), "", TBCompras_Pedido!Observacoes)
        txtRev_serv = IIf(IsNull(TBCompras_Pedido!RevDesenho), "", TBCompras_Pedido!RevDesenho)
        If TBCompras_Pedido!Servico_cliente = True Then Chk_servico_executado_cliente.Value = 1 Else Chk_servico_executado_cliente.Value = 0
        
        txtunservico.ListIndex = -1
        NomeCampo = "a unidade de estoque"
        If IsNull(TBCompras_Pedido!Unidade) = False And TBCompras_Pedido!Unidade <> "" Then txtunservico.Text = TBCompras_Pedido!Unidade
        Cmb_un_com_serv.ListIndex = -1
        NomeCampo = "a unidade comercial"
        If IsNull(TBCompras_Pedido!Unidade_com) = False And TBCompras_Pedido!Unidade_com <> "" Then Cmb_un_com_serv.Text = TBCompras_Pedido!Unidade_com
        cmbfamiliaservico.ListIndex = -1
        NomeCampo = "a família"
        If IsNull(TBCompras_Pedido!Classe) = False And TBCompras_Pedido!Classe <> "" Then cmbfamiliaservico.Text = TBCompras_Pedido!Classe
        
1:
        valor = IIf(Txt_valor_moeda = "", 1, Txt_valor_moeda)
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Projproduto_clientes where Codproduto = " & TBCompras_Pedido!Codproduto & " and idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            If txttipocliente <> "JR" And txttipocliente <> "FR" Then
                txtvlrunitservico = IIf(IsNull(TBFI!PConsumo), "", Format((TBFI!PConsumo / FunVerificaTabelaConversaoUnidade(txtunservico, Cmb_un_com_serv)) / valor, "###,##0.0000000000"))
            Else
                txtvlrunitservico = IIf(IsNull(TBFI!PRevenda), "", Format((TBFI!PRevenda / FunVerificaTabelaConversaoUnidade(txtunservico, Cmb_un_com_serv)) / valor, "###,##0.0000000000"))
            End If
        Else
            If txttipocliente <> "JR" And txttipocliente <> "FR" Then
                txtvlrunitservico = IIf(IsNull(TBCompras_Pedido!PConsumo), "", Format((TBCompras_Pedido!PConsumo / FunVerificaTabelaConversaoUnidade(txtunservico, Cmb_un_com_serv)) / valor, "###,##0.0000000000"))
            Else
                txtvlrunitservico = IIf(IsNull(TBCompras_Pedido!PRevenda), "", Format((TBCompras_Pedido!PRevenda / FunVerificaTabelaConversaoUnidade(txtunservico, Cmb_un_com_serv)) / valor, "###,##0.0000000000"))
            End If
        End If
        TBFI.Close
        
        ProcCarregaDadosCFOPProdServ IIf(IsNull(TBCompras_Pedido!ID_CFOP1), 0, TBCompras_Pedido!ID_CFOP1), False
        
        ProcCarregaComboCodRef cmbreferencia_serv, "P.codproduto = " & TBCompras_Pedido!Codproduto, txtIDcliente, "C", True, True
        
        'Carrega comissão
        Set TBExecucao = CreateObject("adodb.recordset")
        TBExecucao.Open "select * from Vendas_Vendedores where N_Vendedor = " & txtVE, Conexao, adOpenKeyset, adLockOptimistic
        If TBExecucao.EOF = False Then
            If TBExecucao!tipocomissao <> "" And IsNull(TBExecucao!tipocomissao) = False Then
                If TBExecucao!tipocomissao = "V" Then
                    txtComissaoServ = TBExecucao!Comissao
                Else
                    Set TBCFOP = CreateObject("adodb.recordset")
                    If TBExecucao!tipocomissao = "C" Then TBCFOP.Open "select * from Vendas_Vendedores_Clientes where IDVendedor = " & TBExecucao!ID & " and IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao!tipocomissao = "P" Then TBCFOP.Open "select * from Vendas_Vendedores_Produto where IDVendedor = " & TBExecucao!ID & " and IDProduto = " & TBCompras_Pedido!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao!tipocomissao = "CP" Then TBCFOP.Open "select Vendas_Vendedores_Produto.comissao from Vendas_Vendedores_Produto INNER JOIN Vendas_Vendedores_Clientes on Vendas_Vendedores_Produto.idcliente = Vendas_Vendedores_Clientes.Id where Vendas_Vendedores_Produto.IDVendedor = " & TBExecucao!ID & " and Vendas_Vendedores_Produto.IDProduto = " & TBCompras_Pedido!Codproduto & " and Vendas_Vendedores_clientes.IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCFOP.EOF = False Then txtComissaoServ = TBCFOP!Comissao
                    TBCFOP.Close
                End If
            End If
        End If
        TBExecucao.Close
        
        ProcBloqueiaTabsServ
        With txtvlrunitservico
            If TBCompras_Pedido!Valor_bloqueado = True Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
    Else
        USMsgBox ("Não foi encontrado nenhum serviço validado com este código interno ou o mesmo está bloqueado."), vbExclamation, "CAPRIND v5.0"
        ProcLimparServicos True
        ProcLiberaTabsServ
    End If
    TBCompras_Pedido.Close
    ProcCalculaDesconto2
    ProcCalculaValoresServicos
Else
    ProcLimparServicos False
    ProcLiberaTabsServ
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste serviço."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaTabsServ()
On Error GoTo tratar_erro

cmdlistaservicos.TabStop = False
cmbreferencia_serv.TabStop = False
optnovoservico.TabStop = False
OPTnovoservicoman.TabStop = False
txtdescservico.TabStop = False
txtdesccomservico.TabStop = False
cmbfamiliaservico.TabStop = False
txtunservico.TabStop = False
Cmb_un_com_serv.TabStop = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaTabsServ()
On Error GoTo tratar_erro

cmdlistaservicos.TabStop = True
cmbreferencia_serv.TabStop = True
optnovoservico.TabStop = True
OPTnovoservicoman.TabStop = True
txtdescservico.TabStop = True
txtdesccomservico.TabStop = True
cmbfamiliaservico.TabStop = True
txtunservico.TabStop = True
Cmb_un_com_serv.TabStop = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarEscopo()
On Error GoTo tratar_erro

Novo_PI3 = False
Aplic = 11
Compras_Cotacao = False
Compras_Pedido = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoEscopo()
On Error GoTo tratar_erro

If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "escopo de fornecimento", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
txtEscopo = ""
Novo_PI3 = True
Aplic = 11
Compras_Cotacao = False
Compras_Pedido = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

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
If txtStatus <> "ABERTA EM ANALISE" And txtStatus <> "VENDIDA" And txtStatus <> "VENDIDA PARCIAL" And txtStatus <> "FATURADA PARCIAL" Then
    USMsgBox ("Só é permitido alterar o escopo de fornecimento de proposta/pedido com o status aberta em análise, vendida, vendida parcial e faturada parcial."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "o escopo de fornecimento", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM vendas_comercial WHERE cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If Novo_PI3 = True Then
        Evento = "Novo escopo de fornecimento"
        USMsgBox ("Novo escopo de fornecimento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar escopo de fornecimento"
    End If
Else
    TBProduto.AddNew
    USMsgBox ("Novo escopo de fornecimento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo escopo de fornecimento"
End If
TBProduto!Cotacao = txtId
TBProduto!Escopo_fornecimento = txtEscopo
TBProduto.Update
TBProduto.Close

'==================================
Modulo = Formulario
ID_documento = txtId
Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & txtCotacao & " - Rev.: " & txtrevisao
Documento1 = ""
ProcGravaEvento
'==================================

Novo_PI3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdCF_Click()
On Error GoTo tratar_erro

Faturamento = False
Clientes = False
Compras_Pedido = False
Familia_NCM = False
ClassFiscal = False
frmProj_Classificacao_Fiscal.Show 1

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
    
    Sit_REG = 2
    If .Text = "Cliente" Then
        If Vendas_PI = True Then
            ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False
        Else
            ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False
        End If
        frmVendas_LocalizarCliente.Show 1
    ElseIf .Text = "Fornecedor" Then
            If Vendas_PI = True Then
                ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False
            Else
                ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False
            End If
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

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_PI.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_PI.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Vendas_PI.PageCount - 1)
    Else
        TBLISTA_Vendas_PI.AbsolutePage = TBLISTA_Vendas_PI.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Vendas_PI.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_PI1.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_PI1.AbsolutePage = -3 Then
        ProcExibePagina1 (TBLISTA_Vendas_PI1.PageCount - 1)
    Else
        TBLISTA_Vendas_PI1.AbsolutePage = TBLISTA_Vendas_PI1.AbsolutePage - 2
        ProcExibePagina1 (TBLISTA_Vendas_PI1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_PI2.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_PI2.AbsolutePage = -3 Then
        ProcExibePagina2 (TBLISTA_Vendas_PI2.PageCount - 1)
    Else
        TBLISTA_Vendas_PI2.AbsolutePage = TBLISTA_Vendas_PI2.AbsolutePage - 2
        ProcExibePagina2 (TBLISTA_Vendas_PI2.AbsolutePage)
    End If
Else
    ProcExibePagina2 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr2_Click()
On Error GoTo tratar_erro

If txtPagIr2 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas2.Caption, 4))
If Quant <= 1 Or txtPagIr2 > Quant Then Exit Sub
If txtPagIr2.Text >= 1 And txtPagIr2.Text <= Quant Then
    TBLISTA_Vendas_PI2.AbsolutePage = txtPagIr2.Text
    ProcExibePagina2 (TBLISTA_Vendas_PI2.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_PI2.AbsolutePage = 1
ProcExibePagina2 (TBLISTA_Vendas_PI2.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_PI2.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_PI2.AbsolutePage = 1 Then
        ProcExibePagina2 (2)
    Else
        ProcExibePagina2 (TBLISTA_Vendas_PI2.AbsolutePage)
    End If
Else
    ProcExibePagina2 (TBLISTA_Vendas_PI2.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_PI2.AbsolutePage = TBLISTA_Vendas_PI2.PageCount
ProcExibePagina2 (TBLISTA_Vendas_PI2.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr1_Click()
On Error GoTo tratar_erro

If txtPagIr1 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas1.Caption, 4))
If Quant <= 1 Or txtPagIr1 > Quant Then Exit Sub
If txtPagIr1.Text >= 1 And txtPagIr1.Text <= Quant Then
    TBLISTA_Vendas_PI1.AbsolutePage = txtPagIr1.Text
    ProcExibePagina1 (TBLISTA_Vendas_PI1.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_PI1.AbsolutePage = 1
ProcExibePagina1 (TBLISTA_Vendas_PI1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_PI1.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_PI1.AbsolutePage = 1 Then
        ProcExibePagina1 (2)
    Else
        ProcExibePagina1 (TBLISTA_Vendas_PI1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (TBLISTA_Vendas_PI1.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_PI1.AbsolutePage = TBLISTA_Vendas_PI1.PageCount
ProcExibePagina1 (TBLISTA_Vendas_PI1.AbsolutePage)

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
    TBLISTA_Vendas_PI.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Vendas_PI.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_PI.AbsolutePage = 1
ProcExibePagina (TBLISTA_Vendas_PI.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_PI.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_PI.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Vendas_PI.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Vendas_PI.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_PI.AbsolutePage = TBLISTA_Vendas_PI.PageCount
ProcExibePagina (TBLISTA_Vendas_PI.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvar_desconto_Click()
On Error GoTo tratar_erro

'FALTA CONCLUIR

'If Alterar = False Then
'    usMsgbox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'If FunVerifValidacaoRegistro("alterar", txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "o desconto", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
'If Vendas_PI = True Then TextoFiltro = " and (liberacao = 'VENDIDA' or liberacao = 'REVISADA' or liberacao = 'FATURAR' or liberacao = 'FATURAR PARCIAL' or liberacao = 'FATURADO' or liberacao = 'FATURADO PARCIAL' or liberacao = 'CANCELADO')" Else TextoFiltro = ""
'Set TBCotacao = CreateObject("adodb.recordset")
'TBCotacao.Open "Select * from vendas_carteira where cotacao = " & txtid & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
'If TBCotacao.EOF = False Then
'    Contador = TBCotacao.RecordCount
'    Valor = txtTotaldesconto
'    Valor = Format(Valor / Contador, "###,#####0.00000")
'    NovoValor = Replace(Valor, ",", ".")
'
'    Conexao.Execute "UPDATE vendas_carteira Set valordesconto = " & NovoValor & " where cotacao = " & txtid
'    Conexao.Execute "UPDATE vendas_carteira Set Desconto = " & (ValorDesconto * 100) / Preco_unitario & ", Preco_unitario_desconto = Preco_unitario - valordesconto where cotacao = " & txtid
'    ProcGravarTotais txtid
'Else
'    usMsgbox ("Favor cadastrar produto/serviço, antes de salvar o desconto."), vbInformation, "CAPRIND v5.0"
'End If
'TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdtransportadora_Click()
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes where IDCLIENTE = " & txtIDcliente.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    If IsNull(TBClientes!txt_transportadora) = False And TBClientes!txt_transportadora <> "" Then
        cmbtransportadora.Clear
        Select Case TBClientes!Tipo_transp
            Case "C": Cmb_tipo_transp = "Cliente"
            Case "F": Cmb_tipo_transp = "Fornecedor"
            Case "E": Cmb_tipo_transp = "Empresa"
        End Select
        cmbtransportadora.Text = TBClientes!txt_transportadora
    End If
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalentrega_Click()
On Error GoTo tratar_erro

With txtlocal_entrega
    .Clear
    .AddItem ""
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from clientes_entrega where idcliente = " & txtIDcliente.Text & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        Do While TBClientes.EOF = False
            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
            Else
                Endereco = IIf(IsNull(TBClientes!endereco_entrega), "", TBClientes!endereco_entrega)
            End If
            If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
            Else
                Bairro = IIf(IsNull(TBClientes!bairro_entrega), "", TBClientes!bairro_entrega)
            End If
            Endereco1 = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_entrega), "", TBClientes!cidade_entrega) & " - " & IIf(IsNull(TBClientes!uf_entrega), "", TBClientes!uf_entrega) & " - " & IIf(IsNull(TBClientes!cep_entrega), "", TBClientes!cep_entrega)
            ID_entrega = TBClientes!identrega
            
            .AddItem Endereco1
            .ItemData(.NewIndex) = ID_entrega
            TBClientes.MoveNext
        Loop
        txtlocal_entrega = Endereco1
        Txt_ID_entrega = ID_entrega
    End If
    TBClientes.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalcobranca_Click()
On Error GoTo tratar_erro

With txtlocal_cobranca
    .Clear
    .AddItem ""
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from clientes_cobranca where idcliente = " & txtIDcliente.Text & " and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        Do While TBClientes.EOF = False
            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
            Else
                Endereco = IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
            End If
            If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
            Else
                Bairro = IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
            End If
            Endereco1 = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_Cobranca), "", TBClientes!cidade_Cobranca) & " - " & IIf(IsNull(TBClientes!uf_Cobranca), "", TBClientes!uf_Cobranca) & " - " & IIf(IsNull(TBClientes!cep_Cobranca), "", TBClientes!cep_Cobranca)
            ID_Cobranca = TBClientes!idCobranca
            
            .AddItem Endereco1
            .ItemData(.NewIndex) = ID_Cobranca
            TBClientes.MoveNext
        Loop
        txtlocal_cobranca = Endereco1
        Txt_ID_cobranca = ID_Cobranca
    End If
    TBClientes.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelaPI_prod()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtid_produto.Text = 0 Then
    USMsgBox ("Informe o produto antes de cancelar o pedido interno."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente cancelar o pedido interno do produto " & txtNomenclatura & " - Rev." & txtRev_cod & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If Listprod.SelectedItem.ListSubItems.Item(9).Text <> "VENDIDA" And Listprod.SelectedItem.ListSubItems.Item(9).Text <> "VENDIDA PARCIAL" Then
        USMsgBox ("Só é permitido cancelar produto com o status vendida ou vendida parcial."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Vendas_PI = True Then
        If FunVerifValidacaoRegistro("cancelar", txtDtValidacao, "mesmo", "o produto do pedido interno", True) = False Then Exit Sub
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select ID from Producao_pedidos where IDcarteira = " & txtid_produto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            USMsgBox ("Não é permitido cancelar este produto do pedido interno, pois já foi emitida ordem de produção para o mesmo."), vbExclamation, "CAPRIND v5.0"
            TBAbrir.Close
            Exit Sub
        End If
        TBAbrir.Close
    End If
    
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_carteira where Codigo = " & txtid_produto.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        TBVendas!Liberacao = "ABERTA EM ANALISE"
        TBVendas!Datavendas = Null
        TBVendas!PrazoFinal = Null
        TBVendas!Prazo_original = Null
        TBVendas!PCCliente = Null
        txtDatavendas.Text = ""
        TBVendas.Update
        
        ProcExcluirEmpenhos Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, True
    End If
    TBVendas.Close
    
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_PROPOSTA where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_carteira where cotacao = " & TBVendas!Cotacao & " order by datavendas", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!PCCliente) = False Then
                    TBVendas!Datavendas = TBAbrir!Datavendas
                    TBVendas!Tipo = "PRPE"
                Else
                    TBVendas!Datavendas = Null
                    TBVendas!Tipo = "PR"
                End If
                TBAbrir.MoveNext
                TBVendas.Update
            Loop
        End If
        TBAbrir.Close
    End If
    USMsgBox ("Pedido interno do produto " & txtNomenclatura & " - Rev." & txtRev_cod & " cancelado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Cancelar pedido interno do produto"
    ID_documento = txtid_produto
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = "Cód. interno: " & txtNomenclatura
    ProcGravaEvento
    '==================================
    ProcAtualizalistaProdutos (IIf(ReturnNumbersOnly(Left(lblPaginas1.Caption, Len(lblPaginas1.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas1.Caption, Len(lblPaginas1.Caption) - 5))))
    If CodigoLista <> 0 And Listprod.ListItems.Count <> 0 Then
        Listprod.SelectedItem = Listprod.ListItems(CodigoLista)
1:
        Listprod.SetFocus
    End If
    If Vendas_PI = True Then
        If FunAtualizaStatusPropPI(txtId) = True Then
            SSTab1.Tab = 0
            SSTab1_Click (0)
            ProcLimpar
        End If
    Else
        FunAtualizaStatusPropPI txtId
    End If
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelaPI_serv()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtid_servico.Text = 0 Then
    USMsgBox ("Informe o serviço antes de cancelar o pedido interno."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente cancelar o pedido interno do serviço " & txtcodservico & " - Rev." & txtRev_serv & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If ListaServicos.SelectedItem.ListSubItems.Item(9).Text <> "VENDIDA" And ListaServicos.SelectedItem.ListSubItems.Item(6).Text <> "VENDIDA PARCIAL" Then
        USMsgBox ("Só é permitido cancelar serviço com o status vendida ou vendida parcial."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Vendas_PI = True Then
        If FunVerifValidacaoRegistro("cancelar", txtDtValidacao, "mesmo", "o serviço do pedido interno", True) = False Then Exit Sub
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select ID from Producao_pedidos where IDcarteira = " & txtid_servico, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            USMsgBox ("Não é permitido cancelar este serviço do pedido interno, pois já foi emitida ordem de produção para o mesmo."), vbExclamation, "CAPRIND v5.0"
            TBAbrir.Close
            Exit Sub
        End If
        TBAbrir.Close
    End If
    
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_carteira where Codigo = " & txtid_servico.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        TBVendas!Liberacao = "ABERTA EM ANALISE"
        TBVendas!Datavendas = Null
        TBVendas!PrazoFinal = Null
        TBVendas!Prazo_original = Null
        TBVendas!PCCliente = Null
        txtDatavendas.Text = ""
        TBVendas.Update
        
        ProcExcluirEmpenhos Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, True
    End If
    TBVendas.Close
    
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_PROPOSTA where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_carteira where cotacao = " & TBVendas!Cotacao & " order by datavendas", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!PCCliente) = False Then
                    TBVendas!Datavendas = TBAbrir!Datavendas
                    TBVendas!Tipo = "PRPE"
                Else
                    TBVendas!Datavendas = Null
                    TBVendas!Tipo = "PR"
                End If
                TBAbrir.MoveNext
                TBVendas.Update
            Loop
        End If
        TBAbrir.Close
    End If
    USMsgBox ("Pedido interno do serviço " & txtcodservico & " - Rev." & txtRev_serv & " cancelado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Cancelar pedido interno do serviço"
    ID_documento = txtid_servico
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = "Cód. interno: " & txtcodservico
    ProcGravaEvento
    '==================================
    ProcAtualizalistaServicos (IIf(ReturnNumbersOnly(Left(lblPaginas2.Caption, Len(lblPaginas2.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas2.Caption, Len(lblPaginas2.Caption) - 5))))
    If CodigoLista1 <> 0 And ListaServicos.ListItems.Count <> 0 Then
        ListaServicos.SelectedItem = ListaServicos.ListItems(CodigoLista1)
1:
        ListaServicos.SetFocus
    End If
    If Vendas_PI = True Then
        If FunAtualizaStatusPropPI(txtId) = True Then
            SSTab1.Tab = 0
            SSTab1_Click (0)
            ProcLimpar
        End If
    Else
        FunAtualizaStatusPropPI txtId
    End If
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmitirPI_prod()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtid_produto.Text = 0 Then
    USMsgBox ("Informe o produto antes de emitir pedido interno."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente emitir pedido interno deste produto?", vbYesNo) = vbYes Then
    If FunVerifValidarAutomPropPI(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then
        If FunVerificaRegistroValidado("Vendas_proposta", "Cotacao = " & txtId, "mesma", "dessa proposta", "emitir PI", False, False) = False Then Exit Sub
    End If
    If txtIDcliente = "" Or txtIDcliente = "0" Then
        USMsgBox ("Só é permitido emitir pedido interno de proposta com o cliente cadastrado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Listprod.SelectedItem.ListSubItems.Item(9).Text <> "ABERTA EM ANALISE" Then
        USMsgBox ("Só é permitido emitir pedido interno de produto com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_carteira where Codigo = " & txtid_produto.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        'Verif. se o produtos/item está cadastrado
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where desenho = '" & TBVendas!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            USMsgBox ("Não é permitido criar pedido interno deste produto, pois o mesmo precisa estar cadastrado."), vbExclamation, "CAPRIND v5.0"
            TBProduto.Close
            Exit Sub
        End If
        TBProduto.Close
        pc = InputBox("Favor informar o número do pedido de compra do cliente.")
        If pc = "" Then Exit Sub
        
        TextoFiltroUpdate = ""
        Permitido = False
        If FunVerifValidarAutomPropPI(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            Conexao.Execute "UPDATE vendas_proposta Set DtValidacao = '" & Now & "', RespValidacao = '" & pubUsuario & "', DtValidacaoPI = '" & Now & "', RespValidacaoPI = '" & pubUsuario & "' where Cotacao = " & txtId & " and DtValidacao IS NULL"
            Permitido = True
        End If
        
        TBVendas!Liberacao = "VENDIDA"
        TBVendas!Datavendas = Date
        If IsNull(TBVendas!prazofinaldias) = False And TBVendas!prazofinaldias <> "" Then TBVendas!PrazoFinal = FunDefinirPrazoPed(Date + TBVendas!prazofinaldias) Else TBVendas!PrazoFinal = Null
        TBVendas!Prazo_original = TBVendas!PrazoFinal
        TBVendas!PCCliente = pc
        TBVendas.Update
        
        'Muda o tipo da proposta para PRPE
        Conexao.Execute "UPDATE vendas_proposta Set Tipo = 'PRPE', Datavendas = '" & Date & "' where Cotacao = " & txtId
        
        If Permitido = True Then
            QuantSolicitado = TBVendas!Qtde_produzir
            ProcEmpenharProdEstoque Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, TBVendas!Desenho, True, False, TBVendas!Qtde_produzir
            If QuantSolicitado > 0 Then ProcEmpenharProdProduzindo Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, TBVendas!Desenho, TBVendas!PrazoFinal, True
        End If
    End If
    TBVendas.Close
    USMsgBox ("Pedido interno do produto gerado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Emitir pedido interno do produto"
    ID_documento = txtid_produto
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = "Cód. interno: " & txtNomenclatura
    ProcGravaEvento
    '==================================
    ProcAtualizalistaProdutos (IIf(ReturnNumbersOnly(Left(lblPaginas1.Caption, Len(lblPaginas1.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas1.Caption, Len(lblPaginas1.Caption) - 5))))
    If CodigoLista <> 0 And Listprod.ListItems.Count <> 0 Then
        Listprod.SelectedItem = Listprod.ListItems(CodigoLista)
1:
        Listprod.SetFocus
    End If
    FunAtualizaStatusPropPI txtId
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmitirPI_serv()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtid_servico.Text = 0 Then
    USMsgBox ("Informe o serviço antes de emitir o pedido interno."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente emitir pedido interno deste serviço?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifValidarAutomPropPI(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then
        If FunVerificaRegistroValidado("Vendas_proposta", "Cotacao = " & txtId, "mesma", "dessa proposta", "emitir PI", False, False) = False Then Exit Sub
    End If
    If txtIDcliente = "" Or txtIDcliente = "0" Then
        USMsgBox ("Só é permitido emitir pedido interno de proposta com o cliente cadastrado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If ListaServicos.SelectedItem.ListSubItems.Item(9).Text <> "ABERTA EM ANALISE" Then
        USMsgBox ("Só é permitido emitir pedido interno de serviço com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_carteira where Codigo = " & txtid_servico, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        'Verif. se o serviço está cadastrado
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where desenho = '" & TBVendas!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            USMsgBox ("Não é permitido criar pedido interno deste serviço, pois o mesmo precisa estar cadastrado."), vbExclamation, "CAPRIND v5.0"
            TBProduto.Close
            Exit Sub
        End If
        TBProduto.Close
        pc = InputBox("Favor informar o número do pedido de compra do cliente.")
        If pc = "" Then Exit Sub
        
        TextoFiltroUpdate = ""
        Permitido = False
        If FunVerifValidarAutomPropPI(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            Conexao.Execute "UPDATE vendas_proposta Set DtValidacao = '" & Now & "', RespValidacao = '" & pubUsuario & "', DtValidacaoPI = '" & Now & "', RespValidacaoPI = '" & pubUsuario & "' where Cotacao = " & txtId & " and DtValidacao IS NULL"
            Permitido = True
        End If
        
        TBVendas!Liberacao = "VENDIDA"
        TBVendas!Datavendas = Date
        If IsNull(TBVendas!prazofinaldias) = False And TBVendas!prazofinaldias <> "" Then TBVendas!PrazoFinal = FunDefinirPrazoPed(Date + TBVendas!prazofinaldias) Else TBVendas!PrazoFinal = Null
        TBVendas!Prazo_original = TBVendas!PrazoFinal
        TBVendas!PCCliente = pc
        TBVendas.Update
        
        'Muda o tipo da proposta para PRPE
        Conexao.Execute "UPDATE vendas_proposta Set Tipo = 'PRPE', Datavendas = '" & Date & "' where Cotacao = " & txtId
        
        If Permitido = True Then
            QuantSolicitado = TBVendas!Qtde_produzir
            ProcEmpenharProdEstoque Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, TBVendas!Desenho, True, False, TBVendas!Qtde_produzir
            If QuantSolicitado > 0 Then ProcEmpenharProdProduzindo Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, TBVendas!Desenho, TBVendas!PrazoFinal, True
        End If
    End If
    TBVendas.Close
    USMsgBox ("Pedido interno do serviço gerado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Emitir pedido interno do serviço"
    ID_documento = txtid_servico
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = "Cód. interno: " & txtcodservico
    ProcGravaEvento
    '==================================
    ProcAtualizalistaServicos (IIf(ReturnNumbersOnly(Left(lblPaginas2.Caption, Len(lblPaginas2.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas2.Caption, Len(lblPaginas2.Caption) - 5))))
    If CodigoLista1 <> 0 And ListaServicos.ListItems.Count <> 0 Then
        ListaServicos.SelectedItem = ListaServicos.ListItems(CodigoLista1)
1:
        ListaServicos.SetFocus
    End If
    FunAtualizaStatusPropPI txtId
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro

If Vendas_Proposta = True Then
    If txtCotacao = "" Then Exit Sub
    Select Case txtStatus.Text
        Case "VENDIDA": Exit Sub
        Case "VENDIDA PARCIAL": Exit Sub
        Case "REVISADA": Exit Sub
        Case "FATURADA": Exit Sub
        Case "FATURADA PARCIAL": Exit Sub
    End Select
    frmVendas_propostaII_Status.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro
  
frmVendas_PI_MenuImpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarComercial()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus <> "ABERTA EM ANALISE" And txtStatus <> "VENDIDA" And txtStatus <> "VENDIDA PARCIAL" And txtStatus <> "FATURADA PARCIAL" Then
    USMsgBox ("Só é permitido alterar os dados comerciais de " & IIf(Vendas_PI = True, "pedido interno", "proposta") & " com o status aberta em análise, vendida, vendida parcial e faturada parcial."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
'If FunVerifValidacaoRegistro("alterar", txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "os dados comerciais", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
Acao = "salvar"
If cmbMoeda = "" Then
    NomeCampo = "a moeda"
    ProcVerificaAcao
    cmbMoeda.SetFocus
    Exit Sub
End If
valor = IIf(Txt_valor_moeda = "", 0, Txt_valor_moeda)
If valor <= 0 Then
    NomeCampo = "o valor da moeda"
    ProcVerificaAcao
    Txt_valor_moeda.SetFocus
    Exit Sub
End If

If txtTipoTransp(0).Text = "" Then
    NomeCampo = "o tipo da transportadora"
    ProcVerificaAcao
    txtTipoTransp(0).SetFocus
    Exit Sub
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM vendas_comercial WHERE cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dados comerciais"
Else
    TBProduto.AddNew
    USMsgBox ("Dados comerciais " & IIf(Vendas_PI = True, "do pedido interno", "da proposta") & " salvos com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo dados comerciais"
End If
ProcEnviadadosComercial
TBProduto.Update
TBProduto.Close
'==================================
Modulo = Formulario
ID_documento = txtId
Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & txtCotacao & " - Rev.: " & txtrevisao
Documento1 = ""
ProcGravaEvento
'==================================
ProcPuxaDadosComercial
If Vendas_Proposta = True Then ProcSalvarPrevisaoPgto IIf(txttotalproposta = "", 0, txttotalproposta)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosComercial()
On Error GoTo tratar_erro

TBProduto!analize = IIf(txtAnalize = "", Null, txtAnalize)
TBProduto!calculos = txtcalculos.Text
TBProduto!impostos = txtimpostos.Text
TBProduto!condicoes = txtCondicoes.Text
TBProduto!Cotacao = txtId.Text
TBProduto!garantia = txtgarantia.Text
TBProduto!Observacoes = txtObservacoes.Text
TBProduto!reajuste = txtReajuste.Text
TBProduto!transporte = txttransporte.Text
TBProduto!validade = IIf(txtValidade = "", Null, txtValidade)

If txtlocal_entrega <> "" Then
    TBProduto!ID_entrega = Txt_ID_entrega
    TBProduto!Local_entrega = txtlocal_entrega
Else
    TBProduto!ID_entrega = 0
    TBProduto!Local_entrega = Null
End If
If txtlocal_cobranca <> "" Then
    TBProduto!ID_Cobranca = Txt_ID_cobranca
    TBProduto!Local_Cobranca = txtlocal_cobranca
Else
    TBProduto!ID_Cobranca = 0
    TBProduto!Local_Cobranca = Null
End If

'======================================================================================
' Transportadora
'======================================================================================
If txtTransportadora.Text <> "" Then
    Select Case txtTipoTransp(0).Text
        Case "Cliente": TBProduto!Tipo_transp = "C"
        Case "Fornecedor": TBProduto!Tipo_transp = "F"
        Case "Empresa": TBProduto!Tipo_transp = "E"
    End Select
    TBProduto!Transportadora = txtTransportadora.Text
    TBProduto!IdIntTransp = txtidTransportadora.Text
Else
    TBProduto!Tipo_transp = ""
    TBProduto!Transportadora = ""
    TBProduto!IdIntTransp = 0
End If

'======================================================================================
' Redespacho
'======================================================================================
TBProduto!Tipo_Frete = cmb_Tipo_Frete.Text
If txtRedespacho.Text <> "" Then
    Select Case txtTipoTransp(1).Text
        Case "Cliente": TBProduto!Tipo_transp2 = "C"
        Case "Fornecedor": TBProduto!Tipo_transp2 = "F"
        Case "Empresa": TBProduto!Tipo_transp2 = "E"
    End Select
TBProduto!Redespacho = txtRedespacho.Text
Else
TBProduto!Tipo_transp2 = ""
TBProduto!Redespacho = ""
End If

TBProduto!Moeda = IIf(cmbMoeda = "", Null, cmbMoeda)
TBProduto!Valor_moeda = IIf(Txt_valor_moeda = "", Null, Txt_valor_moeda)

'Atualiza valor dos produtos/serviços de acordo com valor da moeda
If IsNull(TBProduto!Valor_moeda) = False Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * FROM vendas_carteira where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            TBLISTA!preco_unitario = TBLISTA!preco_unitario / TBProduto!Valor_moeda
            TBLISTA!preco_unitario_desconto = TBLISTA!preco_unitario_desconto / TBProduto!Valor_moeda
            TBLISTA!preco_lote = TBLISTA!preco_lote / TBProduto!Valor_moeda
            TBLISTA.Update
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
    ProcAtualizalistaProdutos (1)
    ProcAtualizalistaServicos (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarPrevisaoPgto(ValorTotalProposta As Double)
On Error GoTo tratar_erro

Conexao.Execute "DELETE from vendas_proposta_previsaopgto where cotacao = " & txtId
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select condicoes from vendas_comercial where cotacao = " & txtId & " and condicoes IS NOT NULL and condicoes <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    If TBCotacao!condicoes <> " " And IsNumeric(Left(TBCotacao!condicoes, 1)) = True Then
        QtdeSaida = 0
        Contador = 1
        Contador2 = 0
        nPagto = 0
        Valor_Duplicatas = 0
        QtdeSaida = Len(TBCotacao!condicoes)
        TextoCond = ""
        Do While Contador <= QtdeSaida
            If Mid(TBCotacao!condicoes, Contador, 1) = "/" Or Mid(TBCotacao!condicoes, Contador, 1) = "," Or IsNumeric(Mid(TBCotacao!condicoes, Contador, 1)) = True Then
                If TextoCond = "" Then TextoCond = Mid(TBCotacao!condicoes, Contador, 1) Else TextoCond = TextoCond & Mid(TBCotacao!condicoes, Contador, 1)
            End If
            Contador = Contador + 1
        Loop
        
        'Verifica qtde. de parcelas
        Contador = 1
        QtdeSaida = Len(TextoCond)
        Do While Contador <= QtdeSaida
           Do While Mid(TextoCond, Contador, 1) <> "/" And Contador <= QtdeSaida
                Contador2 = Contador2 + 1
                Contador = Contador + 1
            Loop
            nPagto = nPagto + 1
            Contador = Contador + 1
        Loop
        
        'Verifica valor a receber
        mxValorPag = Format(ValorTotalProposta / nPagto, "###,##0.00")
        
        Contador = 1
        Contador3 = 1
        
        Dataini = txt_dataelaborado
        
        Controle = 0
        Do While Contador <= QtdeSaida
            
            Contador2 = 0
            Do While Mid(TBCotacao!condicoes, Contador, 1) <> "/" And Contador <= QtdeSaida
                Contador2 = Contador2 + 1
                Contador = Contador + 1
            Loop
            
            mxCondpag = ReturnNumbersOnly(Mid(TBCotacao!condicoes, Contador3, Contador2))
            Contador3 = Contador3 + Contador2 + 1
                
            Controle = Controle + 1
            DataFim = Format(Dataini + mxCondpag, "DD/MM/YYYY")
            
            If Controle = nPagto Then valor = Format(ValorTotalProposta - Valor_Duplicatas, "###,##0.00") Else valor = mxValorPag
            Valor_Duplicatas = Valor_Duplicatas + mxValorPag
            
            Par1 = Controle
            Par2 = nPagto
            If Len(Par1) = 1 Then
                Par1 = "00" & Par1
            ElseIf Len(Par1) = 2 Then
                    Par1 = "0" & Par1
            End If
            If Len(Par2) = 1 Then
                Par2 = "00" & Par2
            ElseIf Len(Par2) = 2 Then
                Par2 = "0" & Par2
            End If
            Parcela = Par1 & "/" & Par2
            
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from vendas_proposta_previsaopgto", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Cotacao = txtId
            TBGravar!Data = DataFim
            TBGravar!valor = valor
            TBGravar!Parcela = Parcela
            TBGravar.Update
            TBGravar.Close
            
            Contador = Contador + 1
        Loop
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Img_calendario_retorno_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = True
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
Sit_Data = 3
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = True
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

Private Sub ImgCalendario1_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = True
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




Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Validação" And .ListItems.Item(InitFor).SubItems(7) = "Sim" Then
                If Vendas_PI = True And (.ListItems.Item(InitFor).SubItems(6) = "REVISADA" Or .ListItems.Item(InitFor).SubItems(6) = "FATURADA") Or Vendas_Proposta = True And (.ListItems.Item(InitFor).SubItems(6) = "VENDIDA" Or .ListItems.Item(InitFor).SubItems(6) = "REVISADA" Or .ListItems.Item(InitFor).SubItems(6) = "FATURADA" Or .ListItems.Item(InitFor).SubItems(6) = "FATURADA PARCIAL") Then
                    USMsgBox ("Não é permitido cancelar validação, pois " & IIf(Vendas_PI = True, "o pedido", "a proposta") & " está " & .ListItems.Item(InitFor).SubItems(6) & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                If Vendas_PI = True Then
                    'Verifica se já existe empenho no estoque ou na produção
                    Set TBAliquota = CreateObject("adodb.recordset")
                    TBAliquota.Open "Select Codigo from Empresa E INNER JOIN Vendas_proposta VP ON VP.ID_empresa = E.Codigo where VP.Cotacao = " & .ListItems.Item(InitFor) & " and E.Ativar_empenho_autom = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAliquota.EOF = False Then
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select ECEV.ID from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN Vendas_carteira VC ON VC.Codigo = ECEV.ID_carteira where VC.Cotacao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then
                            USMsgBox ("Não é permitido cancelar validação, pois existe empenho no estoque para este pedido."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                        TBProduto.Close
                    End If
                    TBAliquota.Close
                    
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select PP.ID from Producao_pedidos PP INNER JOIN Vendas_carteira VC ON VC.Codigo = PP.IDcarteira where VC.Cotacao = " & .ListItems.Item(InitFor) & " and PP.Expedicao <> 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        USMsgBox ("Não é permitido cancelar validação, pois existe empenho na produção para este pedido."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    TBProduto.Close
                    
'                    'Verifica se já existe necessidade
'                    Set TBProduto = CreateObject("adodb.recordset")
'                    TBProduto.Open "Select PM.Idmateriaprima from Producaomaterial PM INNER JOIN Vendas_carteira VC ON VC.Codigo = PM.ID_carteira where VC.Cotacao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
'                    If TBProduto.EOF = False Then
'                        usMsgbox ("Não é permitido cancelar validação, pois existe necessidade para este pedido."), vbExclamation, "CAPRIND v5.0"
'                        .ListItems.Item(InitFor).Checked = False
'                        Exit Sub
'                    End If
'                    TBProduto.Close
                    
                    'Verifica se foi gerado ordem de faturamento
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select Codigo from vendas_carteira where Cotacao = " & .ListItems.Item(InitFor) & " and (Liberacao = 'FATURAR' or Liberacao = 'FATURAR PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido cancelar validação, pois existe ordem de faturamento aberta para este pedido."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    TBAbrir.Close
                End If
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Codigo from vendas_carteira where Cotacao = " & .ListItems.Item(InitFor) & " and (Liberacao = 'REVISADA' or Liberacao = 'FATURAR' or Liberacao = 'FATURAR PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    USMsgBox ("Não é permitido alterar status, pois existe ordem de faturamento aberta para " & IIf(Vendas_PI = True, "este pedido ou o mesmo está revisado.", "esta proposta ou a mesma está revisada.")), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                TBAbrir.Close
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_CF_Change()
On Error GoTo tratar_erro

ProcCalculaDesconto
ProcAtualizavalores

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_serv()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus <> "ABERTA EM ANALISE" And txtStatus <> "VENDIDA" And txtStatus <> "VENDIDA PARCIAL" And txtStatus <> "FATURADA PARCIAL" Then
    USMsgBox ("Só é permitido criar novo serviço em proposta com o status aberta em análise, vendida parcial ou faturada parcial."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "serviço", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
Novo_PI2 = True
ProcLimparServicos False
Frame1(11).Enabled = True
txtcodservico.SetFocus
ProcLiberaTabsServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId.Text = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
If Vendas_Proposta = True Then
    TBLISTA.Open "Select * from vendas_proposta order by ordenarproposta, cotacao", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * from vendas_proposta WHERE Tipo = 'PE' or tipo = 'PRPE' order by ordenarproposta, cotacao", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBLISTA.BOF = False Then
    TBLISTA.Find ("cotacao = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId.Text = TBLISTA!Cotacao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_proposta where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
        ProcPuxaDados
        ProcLimparTudo
        ProcPuxaDadosComercial
        ProcAtualizalistaProdutos (1)
        ProcAtualizalistaServicos (1)
        ProcCarregaEscopoForn
    Else
        USMsgBox ("Fim dos cadastros de " & IIf(Vendas_Proposta = True, "proposta comercial.", "pedido interno.")), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_PI = False
Novo_PI1 = False
Novo_PI2 = False
Novo_PI3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId.Text = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
If Vendas_Proposta = True Then
    TBLISTA.Open "Select * from vendas_proposta order by ordenarproposta, cotacao", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select * from vendas_proposta WHERE tipo = 'PE' or tipo = 'PRPE' order by ordenarproposta, cotacao", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBLISTA.BOF = False Then
    TBLISTA.Find ("cotacao = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId.Text = TBLISTA!Cotacao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_proposta where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
        ProcPuxaDados
        ProcLimparTudo
        ProcPuxaDadosComercial
        ProcAtualizalistaProdutos (1)
        ProcAtualizalistaServicos (1)
        ProcCarregaEscopoForn
    Else
        USMsgBox ("Fim dos cadastros de " & IIf(Vendas_Proposta = True, "proposta comercial.", "pedido interno.")), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_PI = False
Novo_PI1 = False
Novo_PI2 = False
Novo_PI3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdContato_Click()
On Error GoTo tratar_erro

If txtIDcliente.Text <> "" And txtIDcliente.Text <> "0" Then
    Analise_critica = False
    Telemarketing = False
    Qualidade_PPAP_PSW = False
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    frmVendas_propostaII_contato.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Serv()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaServicos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from vendas_carteira where Codigo = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from vendas_carteira_alteracoes where ID_carteira = " & .ListItems(InitFor) & " and Tipo = '" & IIf(Vendas_PI = True, "VPI", "VPR") & "'"

            '==================================
            Modulo = Formulario
            Evento = "Excluir serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
            
            'Excluir cliente do serviço
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select vendas_proposta.* from (vendas_carteira INNER JOIN vendas_proposta on Vendas_carteira.cotacao = Vendas_proposta.cotacao) INNER JOIN projproduto on vendas_carteira.desenho = projproduto.desenho where Vendas_carteira.desenho = '" & .ListItems(InitFor).SubItems(1) & "' and Vendas_proposta.IDcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = True Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from projproduto where desenho = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select tbl_dados_nota_fiscal.*, tbl_detalhes_nota.codproduto  from tbl_detalhes_nota INNER JOIN tbl_dados_nota_fiscal on tbl_detalhes_nota.id_nota = tbl_dados_nota_fiscal.ID where tbl_detalhes_nota.int_Cod_Produto = '" & .ListItems(InitFor).SubItems(1) & "' and tbl_dados_nota_fiscal.Id_Int_Cliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then Conexao.Execute "DELETE from Projproduto_clientes WHERE codproduto = " & TBItem!Codproduto & " and IDcliente = " & txtIDcliente
                    TBFI.Close
                End If
                TBItem.Close
            End If
            TBOrdem.Close
            
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparServicos False
    ListaServicos.ListItems.Clear
    ProcAtualizalistaServicos (1)
    ListaServicos.SetFocus
    Frame1(11).Enabled = False
    Novo_PI2 = False
    
    If Vendas_PI = True Then
        If FunAtualizaStatusPropPI(txtId) = True Then
            SSTab1.Tab = 0
            SSTab1_Click (0)
            ProcLimpar
        End If
    Else
        FunAtualizaStatusPropPI txtId
    End If
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdlistaservicos_Click()
On Error GoTo tratar_erro

ProcLiberaTabsServ
PI_Produtos = False
PI_Servicos = True
Vendas_Programacao = False
frmVendas_ListaProduto.Show 1
txtcodservico.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_serv()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(11).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If optnovoservico.Value = 0 And OPTnovoservicoman.Value = 0 And txtcodservico = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtcodservico.SetFocus
    Exit Sub
End If
If txtdescservico.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescservico.SetFocus
    Exit Sub
End If
If Vendas_PI = True Then
    With txtpcclienteserv
        If .Text = "" Then
            NomeCampo = "o pedido do cliente"
            ProcVerificaAcao
            .Locked = False
            .TabStop = True
            .SetFocus
            Exit Sub
        End If
    End With
    If IsDate(mskprazoservico) = False Then
        NomeCampo = "o prazo final"
        ProcVerificaAcao
        mskprazoservico.SetFocus
        Exit Sub
    End If
Else
    If txtPrazo_Servico = "" Then
        NomeCampo = "o prazo em dias"
        ProcVerificaAcao
        txtPrazo_Servico.SetFocus
        Exit Sub
    Else
        Valor_Cofins_Prod = txtPrazo_Servico
        If Valor_Cofins_Prod - Int(Valor_Cofins_Prod) > 0 Then
            USMsgBox ("Só é permitido número inteiro no prazo em dias."), vbExclamation, "CAPRIND v5.0"
            txtPrazo_Servico.SetFocus
            Exit Sub
        End If
    End If
End If
If txtdesccomservico.Text = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtdesccomservico.SetFocus
    Exit Sub
End If
If cmbfamiliaservico.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamiliaservico.SetFocus
    Exit Sub
End If
If txtunservico.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    txtunservico.SetFocus
    Exit Sub
End If
If Cmb_un_com_serv.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com_serv.SetFocus
    Exit Sub
End If
valor = IIf(txtqtservico = "", 0, txtqtservico)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtqtservico.SetFocus
    Exit Sub
End If
valor = IIf(txtvlrunitservico = "", 0, txtvlrunitservico)
If valor < 0 Then
    NomeCampo = "o valor unitário"
    ProcVerificaAcao
    txtvlrunitservico.SetFocus
    Exit Sub
End If
If Chk_desc2.Value = 1 Then
    valor = IIf(txtdesconto2 = "", 0, txtdesconto2)
    If valor < 0 Or valor > 100 Then
        NomeCampo = "a porcentagem do desconto"
        ProcVerificaAcao
        txtdesconto2.SetFocus
        Exit Sub
    End If
End If
If Chk_valor_desc2.Value = 1 Then
    valor = IIf(txtvalordesconto2 = "", 0, txtvalordesconto2)
    If valor < 0 Then
        NomeCampo = "o valor do desconto"
        ProcVerificaAcao
        txtvalordesconto2.SetFocus
        Exit Sub
    End If
End If
valor = IIf(txtiss = "", 0, txtiss)
If valor < 0 Then
    NomeCampo = "a porcentagem do ISS"
    ProcVerificaAcao
    txtiss.SetFocus
    Exit Sub
End If

If optnovoservico.Value = 0 And OPTnovoservicoman.Value = 0 And Vendas_PI = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtcodservico.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then
        USMsgBox ("Não foi encontrado nenhum serviço cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtcodservico.SetFocus
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If

If Txt_ID_CFOP_serv <> "" And Txt_ID_CFOP_serv <> "0" And txtuf <> "" And txtuf <> "EX" Then
    If FunVerificaCFOPUF(Txt_ID_CFOP_serv, txtuf) = False Then Exit Sub
End If

'Verifica se ja existe o mesmo serviço na proposta
If Novo_PI2 = True And txtcodservico <> "" And optnovoservico.Value = 0 And OPTnovoservicoman.Value = 0 Then
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from vendas_carteira where cotacao = " & txtId & " and Desenho = '" & txtcodservico & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        USMsgBox ("Já existe um serviço com o código " & txtcodservico & IIf(Vendas_Proposta = True, " nessa proposta", " nesse pedido") & "."), vbExclamation, "CAPRIND v5.0"
        If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            TBCotacao.Close
            Exit Sub
        End If
    End If
End If

'Se for novo serviço
If optnovoservico.Value = 1 Then
    Call Procnovoservico
    If txtReferencia_serv <> "" Then
        cmbreferencia_serv.AddItem txtReferencia_serv
        cmbreferencia_serv = txtReferencia_serv
    End If
    optnovoservico.Value = 0
End If
If OPTnovoservicoman.Value = 1 Then
    If txtcodservico.Text = "" Then
        USMsgBox ("Informe o código interno antes de salvar."), vbExclamation, "CAPRIND v5.0"
        txtcodservico.SetFocus
        Exit Sub
    End If
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtcodservico.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um serviço cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtcodservico.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    Call ProcNovoServicoMan
    If txtReferencia_serv <> "" Then
        cmbreferencia_serv.AddItem txtReferencia_serv
        cmbreferencia_serv = txtReferencia_serv
    End If
    OPTnovoservicoman.Value = 0
End If

If optnovoservico.Value = 0 And OPTnovoservicoman.Value = 0 Then Conexao.Execute "Update projproduto Set RevDesenho = '" & IIf(txtRev_serv.Text = "", 0, txtRev_serv.Text) & "' where desenho = '" & txtcodservico & "'"

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * from vendas_carteira where Codigo = " & txtid_servico, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    If Vendas_PI = True Then
        If TBCotacao!Liberacao <> "VENDIDA" And TBCotacao!Liberacao <> "VENDIDA PARCIAL" And TBCotacao!Liberacao <> "FATURADO PARCIAL" Then
            USMsgBox ("Só é permitido alterar serviço com o status vendido, vendido parcial ou faturado parcial."), vbExclamation, "CAPRIND v5.0"
            TBCotacao.Close
            Exit Sub
        End If
        If TBCotacao!Desenho <> txtcodservico Then
            If FunVerifAltCodQtde(False, True) = False Then Exit Sub
        End If
        valor = txtqtservico
        If TBCotacao!quantidade <> valor Then
            If FunVerifAltCodQtde(False, False) = False Then Exit Sub
        End If
    Else
        If TBCotacao!Liberacao <> "ABERTA EM ANALISE" Then
            USMsgBox ("Só é permitido alterar serviço com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
            TBCotacao.Close
            Exit Sub
        End If
    End If
Else
    
    'Busca o valor cadastrado de limite de credito no cliente
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select txtLimiteCredito from Clientes where idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then LimiteCredito = TBClientes!txtLimiteCredito
    
    TBClientes.Close
    
    If LimiteCredito <> 0 Then
        'Totaliza o limite utilizado pra contas em aberto
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Sum(valor) as LimiteCreditoUtilizadoAberto, idCliente from tbl_contas_receber where idCliente = " & txtIDcliente & " AND Status = 'TÍTULO EM ABERTO' GROUP BY idCliente", Conexao, adOpenKeyset, adLockReadOnly
        If TBContas.EOF = False Then
            LimiteCreditoUtilizadoAberto = TBContas!LimiteCreditoUtilizadoAberto
        Else
            LimiteCreditoUtilizadoAberto = 0
        End If
        TBContas.Close
        
        'Totaliza o limite utilizado pra contas parcial
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Sum(valorpendente) as LimiteCreditoUtilizadoParcial, idCliente from tbl_contas_receber where idCliente = " & txtIDcliente & " AND Status = 'TÍTULO RECEBIDO PARCIAL' and LogSit = 'N' GROUP BY idCliente", Conexao, adOpenKeyset, adLockReadOnly
        If TBContas.EOF = False Then
            LimiteCreditoUtilizadoParcial = TBContas!LimiteCreditoUtilizadoParcial
        Else
            LimiteCreditoUtilizadoParcial = 0
        End If
        TBContas.Close
        
        'Totalliza o saldo do limite utilizado
        
        LimiteCreditoSaldo = LimiteCreditoUtilizadoAberto + LimiteCreditoUtilizadoParcial + txtvalor_total
        LimiteDisponivel = LimiteCredito - (LimiteCreditoUtilizadoAberto + LimiteCreditoUtilizadoParcial)
        
        If LimiteCredito <= LimiteCreditoSaldo Then
            MsgBox ("O limite de credito desse cliente de R$" & LimiteCredito & " foi atingido, o limite disponível é de R$" & LimiteDisponivel & ", não é possivel adicionar esse produto!"), vbCritical + vbOKOnly
            Exit Sub
        End If
    End If
    
    

    TBCotacao.AddNew
    If Vendas_Proposta = True Then
        TBCotacao!Liberacao = "ABERTA EM ANALISE"
    Else
        TBCotacao!Liberacao = "VENDIDA"
        TBCotacao!Datavendas = txtDatavendas_PI
    End If
    TBCotacao!Tem_ordem = False
End If

TBCotacao!Observacoes = txtObs_serv
TBCotacao!Obs_faturamento = Txt_observacoes_fat_serv
TBCotacao!IDAnalise = IDAnalise_servico
TBCotacao!Desenho = txtcodservico.Text
TBCotacao!N_referencia = cmbreferencia_serv
TBCotacao!ID_CFOP = IIf(Txt_ID_CFOP_serv = "", 0, Txt_ID_CFOP_serv)
TBCotacao!quantidade = txtqtservico.Text
TBCotacao!Qtde_produzir = TBCotacao!quantidade / FunVerificaTabelaConversaoUnidade(txtunservico, Cmb_un_com_serv)
TBCotacao!Rev_codinterno = IIf(txtRev_serv.Text = "", "0", txtRev_serv.Text)
TBCotacao!descricao_tecnica = Trim(txtdescservico.Text)
TBCotacao!Descricao = Trim(txtdesccomservico.Text)
If Chk_servico_executado_cliente.Value = 1 Then TBCotacao!Servico_cliente = True Else TBCotacao!Servico_cliente = False
TBCotacao!Familia = cmbfamiliaservico.Text
TBCotacao!Cidade = Cmb_cidade_servico
TBCotacao!Unidade = txtunservico.Text
TBCotacao!Unidade_com = Cmb_un_com_serv
TBCotacao!Tipo = "S"
TBCotacao!Cotacao = txtId.Text
TBCotacao!Desconto = IIf(txtdesconto2.Text = "", 0, txtdesconto2.Text)
TBCotacao!ValorDesconto = IIf(txtvalordesconto2.Text = "", 0, txtvalordesconto2.Text)
TBCotacao!preco_unitario = IIf(txtvlrunitservico = "", 0, txtvlrunitservico)
TBCotacao!preco_unitario_desconto = txtvalorunitariodesc2
TBCotacao!preco_lote = txtvlrtotalservico

'Calcula comissão
Qtde = 0
Qtd = 0
If txtComissaoServ <> "" Then
    Qtde = txtComissaoServ
    Qtd = TBCotacao!preco_lote
    Qtd = (Qtd * Qtde) / 100
    
    TBCotacao!Comissao = txtComissaoServ
    TBCotacao!ValorComissao = Qtd
Else
    TBCotacao!Comissao = 0
    TBCotacao!ValorComissao = 0
End If
If Vendas_PI = True Then
    TBCotacao!PCCliente = txtpcclienteserv.Text
    'Cadastrar automaticamente a alteração da data
    If Novo_PI2 = False And IsNull(TBCotacao!PrazoFinal) = False Then
        If mskprazoservico <> TBCotacao!PrazoFinal Then ProcINSERTINTO "vendas_carteira_alteracoes", "ID_carteira, Data, Responsavel, Data_alteracao, Responsavel_alteracao, Obs, Alteracao_prazo, Padrao, Tipo", "" & txtid_servico & ",'" & Date & "','" & pubUsuario & "','" & Date & "','" & pubUsuario & "', NULL,'ALTERADO O PRAZO DE ENTREGA DE " & Format(TBCotacao!PrazoFinal, "dd/mm/yy") & " PARA " & Format(mskprazoservico, "dd/mm/yy") & "', 'True', 'VPI'"
    End If
    TBCotacao!PrazoFinal = mskprazoservico.Text
    If Novo_PI2 = True And IsNull(TBCotacao!Prazo_original) = True Then TBCotacao!Prazo_original = mskprazo.Text
Else
    'Cadastrar automaticamente a alteração da data
    If Novo_PI2 = False And IsNull(TBCotacao!prazofinaldias) = False Then
        If txtPrazo_Servico <> TBCotacao!prazofinaldias Then ProcINSERTINTO "vendas_carteira_alteracoes", "ID_carteira, Data, Responsavel, Data_alteracao, Responsavel_alteracao, Obs, Alteracao_prazo, Padrao, Tipo", "" & txtid_servico & ",'" & Date & "','" & pubUsuario & "','" & Date & "','" & pubUsuario & "', NULL,'ALTERADO O PRAZO DE ENTREGA DE " & TBCotacao!prazofinaldias & " DIAS PARA " & txtPrazo_Servico & " DIAS', 'True', 'VPR'"
    End If
    TBCotacao!prazofinaldias = IIf(txtPrazo_Servico = "", Null, txtPrazo_Servico)
End If
    
TBCotacao!Caminho_PCCliente = Caminho_PC_serv_PI
If Chk_antecipacao_serv.Value = 1 Then TBCotacao!Antecipacao_fat = True Else TBCotacao!Antecipacao_fat = False
If Chk_faturamento_parcial_serv.Value = 1 Then TBCotacao!Faturamento_parcial = True Else TBCotacao!Faturamento_parcial = False
If Chk_utiliza_mat_consignado_serv.Value = 1 Then TBCotacao!Utiliza_mat_cons = True Else TBCotacao!Utiliza_mat_cons = False

'Impostos
Valor_total = txtvalorunitariodesc2 * txtqtservico

'Empresa
ProcControleImposto IIf(Txt_ID_CFOP_serv = "", 0, Txt_ID_CFOP_serv), IIf(txtIDcliente = "", 0, txtIDcliente)
ProcVerifImpostosEmpresa Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, txtcodservico, IIf(Chk_servico_executado_cliente.Value = 1, True, False), Valor_total, True, TabelaSN_PI, 0
'Novo cálculo simples nacional 2018
TBCotacao!DAS = DAS
If DAS <> 0 Then TBCotacao!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBCotacao!Total_DAS = 0
TBCotacao!PIS_Serv = PIS_Serv
If PIS_Serv <> 0 Then TBCotacao!Total_PIS_serv = Format((Valor_total * PIS_Serv) / 100, "###,##0.00") Else TBCotacao!Total_PIS_serv = 0
TBCotacao!Cofins_Serv = Cofins_Serv
If Cofins_Serv <> 0 Then TBCotacao!Total_Cofins_serv = Format((Valor_total * Cofins_Serv) / 100, "###,##0.00") Else TBCotacao!Total_Cofins_serv = 0
TBCotacao!CSLL_Serv = CSLL_Serv
If CSLL_Serv <> 0 Then TBCotacao!Total_CSLL_serv = Format((Valor_total * CSLL_Serv) / 100, "###,##0.00") Else TBCotacao!Total_CSLL_serv = 0
TBCotacao!ISS = txtiss.Text
If txtvlrISS <> "" Then TBCotacao!VlrISS = Format(txtvlrISS.Text, "###,##0.00") Else TBCotacao!VlrISS = 0
TBCotacao!INSS_Serv = INSS_Serv
If INSS_Serv <> 0 Then TBCotacao!Total_INSS_serv = Format((Valor_total * INSS_Serv) / 100, "###,##0.00") Else TBCotacao!Total_INSS_serv = 0
TBCotacao!IRPJ_Serv = IRPJ_Serv
If IRPJ_Serv <> 0 Then TBCotacao!Total_IRPJ_serv = Format((Valor_total * IRPJ_Serv) / 100, "###,##0.00") Else TBCotacao!Total_IRPJ_serv = 0
TBCotacao!IRRF_Serv = IRRF_Serv
If IRRF_Serv <> 0 Then TBCotacao!Total_IRRF_serv = Format((Valor_total * IRRF_Serv) / 100, "###,##0.00") Else TBCotacao!Total_IRRF_serv = 0
TBCotacao!cpp = CPP_Serv
If CPP_Serv <> 0 Then TBCotacao!Total_CPP = Format((Valor_total * CPP_Serv) / 100, "###,##0.00") Else TBCotacao!Total_CPP = 0

TBCotacao.Update
txtid_servico = TBCotacao!CODIGO

If txtCliente <> "0" Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select codproduto from Projproduto where desenho = '" & txtcodservico & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then ProcAgregarProdutoCli TBFI!Codproduto, txtIDcliente, txttipocliente, txtunservico, Cmb_un_com_serv, IIf(txtvlrunitservico = "", 0, txtvlrunitservico)
    TBFI.Close
End If

'Atualiza valor de venda do serviço
'If txtTipoCliente <> "JR" And txtTipoCliente <> "FR" Then
'    ProcAtualizaValorProdServ False, 0, True, txtvlrunitservico, 0, txtcodservico
'Else
'    ProcAtualizaValorProdServ False, 0, False, 0, txtvlrunitservico, txtcodservico
'End If


ProcAtualizalistaServicos (IIf(ReturnNumbersOnly(Left(lblPaginas2.Caption, Len(lblPaginas2.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas2.Caption, Len(lblPaginas2.Caption) - 5))))
If Novo_PI2 = True Then
    USMsgBox ("Novo serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo serviço"
Else
    Evento = "Alterar serviço"
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    If CodigoLista1 <> 0 And ListaServicos.ListItems.Count <> 0 Then
        ListaServicos.SelectedItem = ListaServicos.ListItems(CodigoLista1)
1:
        ListaServicos.SetFocus
    End If
End If
'==================================
Modulo = Formulario
ID_documento = txtid_servico
Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & txtCotacao & " - Rev.: " & txtrevisao
Documento1 = "Cód. interno: " & txtcodservico
ProcGravaEvento
'==================================
FunAtualizaStatusPropPI txtId
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
Novo_PI2 = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizavalores()
On Error GoTo tratar_erro

ProcValorImposto txtCotacao, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), IIf(txtIDcliente = "", 0, txtIDcliente), txtCliente, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), RegimeEmpresa_PI
ProcControleImposto IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), IIf(txtIDcliente = "", 0, txtIDcliente)
If TemICMS = "SIM" Then
    txtint_icms = IntICMS
Else
    txtint_icms = 0
    txtvalor_icms = "0,00"
End If
If TemIPI = "SIM" Then
    txtInt_ipi = IntIPI
Else
    txtInt_ipi.Text = 0
    txtdbl_valoripi = "0,00"
End If
ProcCalculaValores

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Left(Caption, 34) = "Administrativo - Vendas - Proposta" Or Left(Caption, 17) = "Vendas - Proposta" Then
    Formulario = "Vendas/Proposta comercial"
    Vendas_Proposta = True
    Vendas_PI = False
Else
    Formulario = "Vendas/Pedido interno"
    Vendas_Proposta = False
    Vendas_PI = True
End If
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboMoeda
ProcCarregaComboProduto
ProcCarregaComboServico
ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362P" Then frmVendas_propostaII_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmVendas_propostaII_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza tipo do cliente nas propostas
            Set TBProposta = CreateObject("adodb.recordset")
            TBProposta.Open "Select Idcliente, Tipo_cliente from vendas_proposta order by Idcliente", Conexao, adOpenKeyset, adLockOptimistic
            If TBProposta.EOF = False Then
                Do While TBProposta.EOF = False
                    Set TBClientes = CreateObject("adodb.recordset")
                    TBClientes.Open "Select Tipo from clientes where IDCliente = " & TBProposta!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        TBProposta!Tipo_cliente = TBClientes!Tipo
                        TBProposta.Update
                    End If
                    TBClientes.Close
                    TBProposta.MoveNext
                Loop
            End If
        End If
            
        If .Chk2.Value = 1 Then
            'Atualiza totais das propostas (valores e impostos dos produtos e serviços)
            Set TBProposta = CreateObject("adodb.recordset")
            TBProposta.Open "Select * from vendas_proposta order by cotacao", Conexao, adOpenKeyset, adLockOptimistic
            If TBProposta.EOF = False Then
                Do While TBProposta.EOF = False
                    txtIDcliente = IIf(IsNull(TBProposta!IDCliente), "", TBProposta!IDCliente)
                    txtCliente = IIf(IsNull(TBProposta!Cliente), "", TBProposta!Cliente)
                    Set TBCarteira = CreateObject("adodb.recordset")
                    TBCarteira.Open "Select * from vendas_carteira where cotacao = " & TBProposta!Cotacao & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCarteira.EOF = False Then
                        Do While TBCarteira.EOF = False
                             If IsNull(TBCarteira!preco_unitario_desconto) = False And (TBCarteira!preco_unitario_desconto) <> 0 Then
                                TBCarteira!preco_lote = TBCarteira!preco_unitario_desconto * TBCarteira!quantidade
                            Else
                                TBCarteira!preco_lote = TBCarteira!preco_unitario * TBCarteira!quantidade
                            End If
                            
                            'Produtos
                            If TBCarteira!Tipo = "P" Then
                                'Impostos
                                Valor_total = IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote)
                                Valor_IPI = IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi)
                                
                                'Empresa
                                TemICMS = "NÃO"
                                TemIPI = "NÃO"
                                TemPIS = False
                                TemCOFINS = False
                                SomarIPI = "NÃO"
                                DestacaImpostos = "NÃO"
                                Set TBMateriaprima = CreateObject("adodb.recordset")
                                TBMateriaprima.Open "Select * FROM tbl_NaturezaOperacao WHERE IDCountCfop = " & IIf(IsNull(TBCarteira!ID_CFOP), 0, TBCarteira!ID_CFOP), Conexao, adOpenKeyset, adLockOptimistic
                                If TBMateriaprima.EOF = False Then
                                    TemICMS = TBMateriaprima!Txt_ICMS
                                    TemIPI = TBMateriaprima!txt_IPI
                                    TemPIS = TBMateriaprima!TemPIS
                                    TemCOFINS = TBMateriaprima!TemCOFINS
                                    SomarIPI = TBMateriaprima!txt_Somar
                                    If TBMateriaprima!Retem = True Then DestacaImpostos = "SIM" Else DestacaImpostos = "NÃO"
                                End If
                                TBMateriaprima.Close
                                
                                ProcVerifImpostosEmpresa TBProposta!ID_empresa, TBProposta!retorno, "", False, 0, False, IIf(IsNull(TBProposta!Tabela_SN), 0, TBProposta!Tabela_SN), 0
                                
                                TBCarteira!PIS_Prod = PIS_Prod
                                If PIS_Prod <> 0 Then TBCarteira!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00") Else TBCarteira!Total_PIS_prod = 0
                                TBCarteira!Cofins_Prod = Cofins_Prod
                                If Cofins_Prod <> 0 Then TBCarteira!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00") Else TBCarteira!Total_Cofins_prod = 0
                                TBCarteira!CSLL_Prod = CSLL_Prod
                                If CSLL_Prod <> 0 Then TBCarteira!Total_CSLL_prod = Format((Valor_total * CSLL_Prod) / 100, "###,##0.00") Else TBCarteira!Total_CSLL_prod = 0
                                TBCarteira!IRPJ_Prod = IRPJ_Prod
                                If IRPJ_Prod <> 0 Then TBCarteira!Total_IRPJ_prod = Format((Valor_total * IRPJ_Prod) / 100, "###,##0.00") Else TBCarteira!Total_IRPJ_prod = 0
                                
                                If IsNull(TBCarteira!ID_CF) = False Then
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select * from tbl_classificacaofiscal where Idclass = " & TBCarteira!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBFI.EOF = False Then
                                        'Verifica se a CF tem retenção de PIS/Cofins, destaca PIS/Cofins e grava no produto
                                        If TBFI!Retem_PIS_Cofins = True Then
                                            TBCarteira!Valor_Retencao_PIS = Format((Valor_total * IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)) / 100, "###,##0.00")
                                            TBCarteira!Valor_Retencao_Cofins = Format((Valor_total * IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)) / 100, "###,##0.00")
                                            
'                                            If Regime <> 1 Then
'                                                PIS_Prod = IIf(IsNull(TBFI!PIS_destaca), 0, TBFI!PIS_destaca)
'                                                Cofins_Prod = IIf(IsNull(TBFI!Cofins_destaca), 0, TBFI!Cofins_destaca)
'                                                If PIS_Prod <> 0 Then
'                                                    TBCarteira!PIS_Prod = PIS_Prod
'                                                    TBCarteira!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00")
'                                                End If
'                                                If Cofins_Prod <> 0 Then
'                                                    TBCarteira!Cofins_Prod = Cofins_Prod
'                                                    TBCarteira!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00")
'                                                End If
'                                            End If
                                            
                                            Valor_total = 0
                                            Valor_IPI = 0
                                        End If
                                    End If
                                    TBFI.Close
                                End If
                                                       
                                If IsNull(TBCarteira!int_IPI) = True Then TBCarteira!int_IPI = 0
                                If IsNull(TBCarteira!IntICMS) = True Then TBCarteira!IntICMS = 0
                                
                                'Se tem IPI
                                If TemIPI = "SIM" Then
                                    Set TBFIltro = CreateObject("adodb.recordset")
                                    TBFIltro.Open "Select * from Clientes_Impostos where IDCliente = " & TBProposta!IDCliente & " and ID_CF = " & IIf(IsNull(TBCarteira!ID_CF), 0, TBCarteira!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
                                    If TBFIltro.EOF = False Then
                                        VlrIPI = IIf(IsNull(TBCarteira!preco_unitario_desconto), 0, TBCarteira!preco_unitario_desconto)
                                        If TBFIltro!PorcentagemIPI <> 0 Then VlrIPI = VlrIPI / TBFIltro!PorcentagemIPI
                                        TBCarteira!dbl_valoripi = Format((VlrIPI - IIf(IsNull(TBCarteira!preco_unitario_desconto), 0, TBCarteira!preco_unitario_desconto)) * IIf(IsNull(TBCarteira!quantidade), 0, TBCarteira!quantidade), "###,##0.00")
                                    Else
                                        TBCarteira!dbl_valoripi = Format((TBCarteira!preco_lote * IIf(IsNull(TBCarteira!int_IPI), 0, TBCarteira!int_IPI)) / 100, "###,##0.00")
                                    End If
                                    TBFIltro.Close
                                Else
                                    TBCarteira!int_IPI = 0
                                    TBCarteira!dbl_valoripi = 0
                                End If
                                
                                ProcCalculaBC TBProposta!ID_empresa, IIf(IsNull(TBProposta!CFOP), 0, TBProposta!CFOP), 0, IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote), IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi), SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBCarteira!txt_CST), 0, TBCarteira!txt_CST), "P", 0, ""
                                TBCarteira!dbl_Valor_ICMS = Format((BC * IIf(IsNull(TBCarteira!IntICMS), 0, TBCarteira!IntICMS)) / 100, "###,##0.00")
                            Else
                                'Impostos
                                Valor_total = IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote)
                                
                                'Empresa
                                DestacaImpostos = "NÃO"
                                Set TBMateriaprima = CreateObject("adodb.recordset")
                                TBMateriaprima.Open "Select * FROM tbl_NaturezaOperacao WHERE IDCountCfop = " & IIf(IsNull(TBCarteira!ID_CFOP), 0, TBCarteira!ID_CFOP), Conexao, adOpenKeyset, adLockOptimistic
                                If TBMateriaprima.EOF = False Then
                                    If TBMateriaprima!Retem = True Then DestacaImpostos = "SIM" Else DestacaImpostos = "NÃO"
                                End If
                                TBMateriaprima.Close
                                ProcVerifImpostosEmpresa TBProposta!ID_empresa, False, TBCarteira!Desenho, TBCarteira!Servico_cliente, Valor_total, True, TabelaSN_PI, 0
                                
                                TBCarteira!PIS_Serv = PIS_Serv
                                If PIS_Serv <> 0 Then TBCarteira!Total_PIS_serv = Format((Valor_total * PIS_Serv) / 100, "###,##0.00") Else TBCarteira!Total_PIS_serv = 0
                                TBCarteira!Cofins_Serv = Cofins_Serv
                                If Cofins_Serv <> 0 Then TBCarteira!Total_Cofins_serv = Format((Valor_total * Cofins_Serv) / 100, "###,##0.00") Else TBCarteira!Total_Cofins_serv = 0
                                TBCarteira!CSLL_Serv = CSLL_Serv
                                If CSLL_Serv <> 0 Then TBCarteira!Total_CSLL_serv = Format((Valor_total * CSLL_Serv) / 100, "###,##0.00") Else TBCarteira!Total_CSLL_serv = 0
                                If IsNull(TBCarteira!ISS) = True Then TBCarteira!ISS = ISS
                                If IsNull(TBCarteira!VlrISS) = True Then TBCarteira!VlrISS = Format((IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote) * ISS) / 100, "###,##0.00")
                                TBCarteira!INSS_Serv = INSS_Serv
                                If INSS_Serv <> 0 Then TBCarteira!Total_INSS_serv = Format((Valor_total * INSS_Serv) / 100, "###,##0.00") Else TBCarteira!Total_INSS_serv = 0
                                TBCarteira!IRPJ_Serv = IRPJ_Serv
                                If IRPJ_Serv <> 0 Then TBCarteira!Total_IRPJ_serv = Format((Valor_total * IRPJ_Serv) / 100, "###,##0.00") Else TBCarteira!Total_IRPJ_serv = 0
                                TBCarteira!IRRF_Serv = IRRF_Serv
                                If IRRF_Serv <> 0 Then TBCarteira!Total_IRRF_serv = Format((Valor_total * IRRF_Serv) / 100, "###,##0.00") Else TBCarteira!Total_IRRF_serv = 0
                            End If
                            TBCarteira.Update
                            TBCarteira.MoveNext
                        Loop
                    End If
                    TBCarteira.Close
                    txtCotacao = TBProposta!Ncotacao
                    txtrevisao = TBProposta!Revisao
                    txtId = TBProposta!Cotacao
                    If IsNull(TBProposta!Tipo_cliente) = False And IsNull(TBProposta!UF) = False Then
                        If IsNull(TBProposta!Tipo_cliente) = False And TBProposta!Tipo_cliente <> "" Then txttipocliente = TBProposta!Tipo_cliente
                        If IsNull(TBProposta!UF) = False And TBProposta!UF <> "" Then
                            With txtuf
                                .Clear
                                .AddItem TBProposta!UF
                                .Text = TBProposta!UF
                            End With
                        End If
                        ProcAtualizalistaProdutos (1)
                        ProcAtualizalistaServicos (1)
                    End If
                    TBProposta.MoveNext
                Loop
            End If
            TBProposta.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Clientes nos produtos/serviços
            Set TBProposta = CreateObject("adodb.recordset")
            TBProposta.Open "Select vendas_proposta.Tipo_cliente, vendas_proposta.IDcliente, vendas_carteira.* from vendas_proposta INNER JOIN vendas_carteira ON vendas_proposta.cotacao = vendas_carteira.Cotacao order by vendas_proposta.ordenarproposta, vendas_proposta.cotacao", Conexao, adOpenKeyset, adLockOptimistic
            If TBProposta.EOF = False Then
                Do While TBProposta.EOF = False
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select Projproduto_clientes.* from Projproduto_clientes INNER JOIN Projproduto on Projproduto_clientes.codproduto = Projproduto.codproduto where Projproduto.desenho = '" & TBProposta!Desenho & "' and Projproduto_clientes.IDcliente = " & TBProposta!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = True Then
                        TBOrdem.AddNew
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select codproduto from Projproduto where desenho = '" & TBProposta!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then TBOrdem!Codproduto = TBFI!Codproduto
                        TBOrdem!IDCliente = TBProposta!IDCliente
                    End If
                    If TBProposta!Tipo_cliente <> "JR" And TBProposta!Tipo_cliente <> "FR" Then TBOrdem!PConsumo = TBProposta!preco_unitario Else TBOrdem!PRevenda = TBProposta!preco_unitario
                    TBOrdem.Update
                    TBOrdem.Close
                    TBProposta.MoveNext
                Loop
            End If
            TBProposta.Close
        End If
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
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

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_PI = True Then
    If USMsgBox("A proposta comercial ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_PI = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PI1 = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_prod
        If Novo_PI1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PI2 = True Then
    If USMsgBox("O serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_serv
        If Novo_PI2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_PI3 = True Then
    If USMsgBox("O escopo de fornecimento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarEscopo
        If Novo_PI3 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_PI = False
Novo_PI1 = False
Novo_PI2 = False
Novo_PI3 = False
Unload Me

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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Validação" And .ListItems.Item(InitFor).SubItems(7) = "Sim" Then
                    If Vendas_PI = True And (.ListItems.Item(InitFor).SubItems(6) = "REVISADA" Or .ListItems.Item(InitFor).SubItems(6) = "FATURADA") Or Vendas_Proposta = True And (.ListItems.Item(InitFor).SubItems(6) = "VENDIDA" Or .ListItems.Item(InitFor).SubItems(6) = "REVISADA" Or .ListItems.Item(InitFor).SubItems(6) = "FATURADA" Or .ListItems.Item(InitFor).SubItems(6) = "FATURADA PARCIAL") Then GoTo Proximo
                    
                    If Vendas_PI = True Then
                        'Verifica se já existe empenho no estoque ou na produção
                        Set TBAliquota = CreateObject("adodb.recordset")
                        TBAliquota.Open "Select Codigo from Empresa E INNER JOIN Vendas_proposta VP ON VP.ID_empresa = E.Codigo where VP.Cotacao = " & .ListItems.Item(InitFor) & " and E.Ativar_empenho_autom = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAliquota.EOF = False Then
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select ECEV.ID from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN Vendas_carteira VC ON VC.Codigo = ECEV.ID_carteira where VC.Cotacao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then GoTo Proximo
                            TBProduto.Close
                        End If
                        TBAliquota.Close
                        
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select PP.ID from Producao_pedidos PP INNER JOIN Vendas_carteira VC ON VC.Codigo = PP.IDcarteira where VC.Cotacao = " & .ListItems.Item(InitFor) & " and PP.Expedicao <> 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then GoTo Proximo
                        TBProduto.Close
                        
'                        'Verifica se já existe necessidade
'                        Set TBProduto = CreateObject("adodb.recordset")
'                        TBProduto.Open "Select PM.Idmateriaprima from Producaomaterial PM INNER JOIN Vendas_carteira VC ON VC.Codigo = PM.ID_carteira where VC.Cotacao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
'                        If TBProduto.EOF = False Then GoTo Proximo
'                        TBProduto.Close
                        
                        'Verifica se foi gerado ordem de faturamento
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select Codigo from vendas_carteira where Cotacao = " & .ListItems.Item(InitFor) & " and (Liberacao = 'FATURAR' or Liberacao = 'FATURAR PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then GoTo Proximo
                        TBAbrir.Close
                    End If
                Else
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select Codigo from vendas_carteira where Cotacao = " & .ListItems.Item(InitFor) & " and (Liberacao = 'REVISADA' or Liberacao = 'FATURAR' or Liberacao = 'FATURAR PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then GoTo Proximo
                    TBAbrir.Close
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select VP.*, CL.CPF_CNPJ as CNPJ_CPF, CL.CEP as CEP, CL.RG_IE from vendas_proposta VP inner join Clientes CL on VP.IDcliente = CL.IDCliente where cotacao ="
'Debug.print StrSql & Lista.SelectedItem

TBAbrir.Open StrSql & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then
    Novo_PI = False
    ProcPuxaDados
    ProcPuxaTotais
    CodigoLista2 = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaServicos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaServicos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If cmbOpcao_lista_prod = "Excluir" Then
                    If txtDtValidacao <> "" Then GoTo Proximo
                    If Vendas_Proposta = True Then
                        If .ListItems(InitFor).SubItems(9) <> "ABERTA EM ANALISE" Then GoTo Proximo
                    Else
                        ProcVerificaRegistroUtilizadoSemMsg "vendas_carteira", "codigo = " & .ListItems(InitFor) & " and liberacao <> 'VENDIDA' and liberacao <> 'VENDIDA PARCIAL'"
                        If Permitido = False Then GoTo Proximo
                    End If
                Else
                    If .ListItems(InitFor).SubItems(9) = "REVISADA" Or .ListItems(InitFor).SubItems(9) = "FATURAR" Or .ListItems(InitFor).SubItems(9) = "FATURAR PARCIAL" Then GoTo Proximo
                    
                    'Se o serviço foi faturado total ele não deixa alterar o status
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "select codigo from vendas_carteira where codigo = " & .ListItems(InitFor) & " and liberacao = 'FATURADO' and DtCancelado IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        TBProduto.Close
                        GoTo Proximo
                    End If
                    TBProduto.Close
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaServicos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaservicos_DblClick()
On Error GoTo tratar_erro

With ListaServicos
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturadaProdServ .SelectedItem, .SelectedItem.ListSubItems(1), False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaservicos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaServicos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If cmbOpcao_lista_prod = "Excluir" Then
                If txtDtValidacao <> "" Then
                    USMsgBox ("Não é permitido excluir este serviço, pois " & IIf(Vendas_Proposta = True, "a proposta está validada", "o pedido está validado") & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If Vendas_Proposta = True Then
                    If .ListItems(InitFor).SubItems(9) <> "ABERTA EM ANALISE" Then
                        USMsgBox ("Só é permitido excluir serviço com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                Else
                    If .ListItems(InitFor).SubItems(9) <> "VENDIDA" And .ListItems(InitFor).SubItems(9) <> "VENDIDA PARCIAL" Then
                        USMsgBox ("Só é permitido excluir serviço com o status " & .ListItems(InitFor).SubItems(9) & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
            Else
                If .ListItems(InitFor).SubItems(9) = "REVISADA" Or .ListItems(InitFor).SubItems(9) = "FATURAR" Or .ListItems(InitFor).SubItems(9) = "FATURAR PARCIAL" Then
                    USMsgBox ("Não é permitido alterar o status deste serviço, pois o mesmo está vinculado a ordem de faturamento ou está revisado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                'Se o serviço foi faturado total ele não deixa alterar o status
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "select codigo from vendas_carteira where codigo = " & .ListItems(InitFor) & " and liberacao = 'FATURADO' and DtCancelado IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then
                    USMsgBox ("Não é permitido alterar o status deste serviço, pois o mesmo está faturado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBProduto.Close
                    Exit Sub
                End If
                TBProduto.Close
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaservicos_itemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaServicos.ListItems.Count = 0 Then Exit Sub
Novo_PI2 = False
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from vendas_carteira where Codigo = " & ListaServicos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimparServicos False
    Frame1(11).Enabled = True
    ProcpuxadadoslistaServicos
    CodigoLista1 = ListaServicos.SelectedItem.index
End If
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
                If cmbOpcao_lista_prod = "Excluir" Then
                    If txtDtValidacao <> "" Then GoTo Proximo
                    If Vendas_Proposta = True Then
                        If .ListItems(InitFor).SubItems(9) <> "ABERTA EM ANALISE" Then GoTo Proximo
                    Else
                        If cmbOpcao_lista_prod = "Excluir" Then
                            If .ListItems(InitFor).SubItems(9) <> "VENDIDA" And .ListItems(InitFor).SubItems(9) <> "VENDIDA PARCIAL" Then GoTo Proximo
                            'Verifica se o produto esta amarrado a uma programação
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "select codigo from vendas_carteira where codigo = " & .ListItems(InitFor) & " and ID_programacao <> 0", Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                TBProduto.Close
                                GoTo Proximo
                            End If
                            TBProduto.Close
                        End If
                    End If
                Else
                    If .ListItems(InitFor).SubItems(9) = "REVISADA" Or .ListItems(InitFor).SubItems(9) = "FATURAR" Or .ListItems(InitFor).SubItems(9) = "FATURAR PARCIAL" Then GoTo Proximo
                    
                    'Se o produto foi faturado total ele não deixa alterar o status
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "select codigo from vendas_carteira where codigo = " & .ListItems(InitFor) & " and liberacao = 'FATURADO' and DtCancelado IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        TBProduto.Close
                        GoTo Proximo
                    End If
                    TBProduto.Close
                End If
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
    ProcVerifQtdeFaturadaProdServ .SelectedItem, .SelectedItem.ListSubItems(1), False
End With

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
            If cmbOpcao_lista_prod = "Excluir" Then
                If txtDtValidacao <> "" Then
                    USMsgBox ("Não é permitido excluir este produto, pois " & IIf(Vendas_Proposta = True, "a proposta está validada", "o pedido está validado") & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If Vendas_Proposta = True Then
                    If .ListItems(InitFor).SubItems(9) <> "ABERTA EM ANALISE" Then
                        USMsgBox ("Só é permitido excluir produto com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                Else
                    If .ListItems(InitFor).SubItems(9) <> "VENDIDA" And .ListItems(InitFor).SubItems(9) <> "VENDIDA PARCIAL" Then
                        USMsgBox ("Não é permitido excluir produto com o status " & .ListItems(InitFor).SubItems(9) & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    'Verifica se o produto esta amarrado a uma programação
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "select codigo from vendas_carteira where codigo = " & .ListItems(InitFor) & " and ID_programacao <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        USMsgBox ("Não é permitido excluir este produto, pois o mesmo está vinculado a uma programação."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBProduto.Close
                        Exit Sub
                    End If
                    TBProduto.Close
                End If
            Else
                If .ListItems(InitFor).SubItems(9) = "REVISADA" Or .ListItems(InitFor).SubItems(9) = "FATURAR" Or .ListItems(InitFor).SubItems(9) = "FATURAR PARCIAL" Then
                    USMsgBox ("Não é permitido alterar o status deste produto, pois o mesmo está vinculado a ordem de faturamento ou está revisado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                'Se o produto foi faturado total ele não deixa alterar o status
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "select codigo from vendas_carteira where codigo = " & .ListItems(InitFor) & " and liberacao = 'FATURADO' and DtCancelado IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = False Then
                    USMsgBox ("Não é permitido alterar o status deste produto, pois o mesmo está faturado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBProduto.Close
                    Exit Sub
                End If
                TBProduto.Close
            End If
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
ProcCarregaCST
Novo_PI1 = False
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from vendas_carteira where Codigo = " & Listprod.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimparProdutos False
    Frame1(10).Enabled = True
    Frame1(12).Enabled = True
    ProcPuxaDadosLista
    CodigoLista = Listprod.SelectedItem.index
    Id_Item = Lista.SelectedItem
End If
ProcLiberaTabsProd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTnovo_Click()
On Error GoTo tratar_erro

If OPTnovo.Value = 1 Then
    ProcLiberaTabsProd
    OPTnovoman.Value = 0
    Procliberacampos
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTnovoman_Click()
On Error GoTo tratar_erro

If OPTnovoman.Value = 1 Then
    ProcLiberaTabsProd
    OPTnovo.Value = 0
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

Private Sub optnovoservico_Click()
On Error GoTo tratar_erro

If optnovoservico.Value = 1 Then
    ProcLiberaTabsServ
    OPTnovoservicoman.Value = 0
    ProcLiberaCamposSev
Else
    ProcBloqueiaCamposServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OPTnovoservicoman_Click()
On Error GoTo tratar_erro

If OPTnovoservicoman.Value = 1 Then
    ProcLiberaTabsServ
    optnovoservico.Value = 0
    ProcLiberaCamposSev
    USMsgBox ("Informe o código interno do serviço."), vbInformation, "CAPRIND v5.0"
    txtcodservico.Text = ""
    txtcodservico.SetFocus
Else
    ProcBloqueiaCamposServ
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procliberacampos()
On Error GoTo tratar_erro

If txtDtValidacao <> "" Then Exit Sub
With txtRev_cod
    .Locked = False
    .TabStop = True
End With
With txtdesctecnica
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
With txtvalorunitario
    .Locked = False
    .TabStop = True
End With
If OPTnovo.Value = 1 Or OPTnovoman.Value = 1 Then
    cmbReferencia.Visible = False
    txtreferencia.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With txtRev_cod
    .Locked = True
    .TabStop = False
End With
With txtdesctecnica
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With
If cmbun <> "KG" And cmbun <> "MM" And cmbun <> "MT" And cmbun <> "PC" And cmbun <> "PÇ" Then
    With Cmb_un_com
        .Locked = True
        .TabStop = False
    End With
End If
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

Sub ProcLiberaCamposSev()
On Error GoTo tratar_erro

With txtRev_serv
    .Locked = False
    .TabStop = True
End With
With txtdescservico
    .Locked = False
    .TabStop = True
End With
With txtunservico
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com_serv
    .Locked = False
    .TabStop = True
End With
With cmbfamiliaservico
    .Locked = False
    .TabStop = True
End With
With txtvlrunitservico
    .Locked = False
    .TabStop = True
End With
If optnovoservico.Value = 1 Or OPTnovoservicoman.Value = 1 Then
    cmbreferencia_serv.Visible = False
    txtReferencia_serv.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCamposServ()
On Error GoTo tratar_erro

With txtRev_serv
    .Locked = True
    .TabStop = False
End With
With txtdescservico
    .Locked = True
    .TabStop = False
End With
With txtunservico
    .Locked = True
    .TabStop = False
End With
If txtunservico <> "KG" And txtunservico <> "MM" And txtunservico <> "MT" And txtunservico <> "PC" And txtunservico <> "PÇ" Then
    With Cmb_un_com_serv
        .Locked = True
        .TabStop = False
    End With
End If
With cmbfamiliaservico
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

Private Sub Txt_CFOP_prod_Change()
On Error GoTo tratar_erro

ProcCarregaCST

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaCST()
On Error GoTo tratar_erro

If Txt_ID_CFOP_prod = "" Then Exit Sub
Cmb_CST_ICMS.Clear
'Cmb_CST_ICMS.AddItem ""
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

Private Sub Txt_ID_CFOP_prod_Change()
On Error GoTo tratar_erro

If Txt_ID_CFOP_prod <> "" Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_NaturezaOperacao where IdCountCfop = " & Txt_ID_CFOP_prod & "", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then If TBFI!retorno = True Then chkRetorno.Value = 1 Else chkRetorno.Value = 0
    TBFI.Close
End If



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

Private Sub txtcodservico_Change()
On Error GoTo tratar_erro

If optnovoservico.Value = 0 And OPTnovoservicoman.Value = 0 Then
    ProcLiberaTabsServ
    ProcLimparServicos True
    If txtcodservico <> "" Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from projproduto where desenho = '" & txtcodservico & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            ProcBloqueiaCamposServ
        Else
            ProcLiberaCamposSev
        End If
        TBFI.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcodservico_LostFocus()
On Error GoTo tratar_erro

If txtcodservico <> "" And optnovoservico.Value = 0 And OPTnovoservicoman.Value = 0 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtcodservico & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        ProcBloqueiaCamposServ
    Else
        ProcLiberaCamposSev
    End If
    TBProduto.Close
Else
    ProcLiberaCamposSev
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtcomissao_Change()
On Error GoTo tratar_erro

If txtComissao <> "" Then
    VerifNumero = txtComissao
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissao = ""
        txtComissao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoServ_Change()
On Error GoTo tratar_erro

If txtComissaoServ <> "" Then
    VerifNumero = txtComissaoServ
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissaoServ = ""
        txtComissaoServ.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComprimento_Change()
On Error GoTo tratar_erro

If txtComprimento <> "" Then
    VerifNumero = txtComprimento
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComprimento = ""
        txtComprimento.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub txtCotacao_Change()
On Error GoTo tratar_erro

If Novo_PI = True Then
VerifCodigo:
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from vendas_proposta where ncotacao = '" & txtCotacao & "' and Cotacao <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Cotacao = Left(txtCotacao, Len(txtCotacao) - 3) + 1
        Ano = Right(Year(Date), 2)
        Select Case Len(Cotacao)
            Case 1: NumeroCotacao = "000" & Cotacao & "/" & Ano
            Case 2: NumeroCotacao = "00" & Cotacao & "/" & Ano
            Case 3: NumeroCotacao = "0" & Cotacao & "/" & Ano
            Case 4: NumeroCotacao = Cotacao & "/" & Ano
            Case 5: NumeroCotacao = Cotacao & "/" & Ano
        End Select
        txtCotacao = NumeroCotacao
        GoTo VerifCodigo
    End If
    TBFI.Close
End If

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
    ProcCalculaDesconto
ElseIf Chk_valor_desc.Value = 0 Then
        txtDesconto = ""
        txtvalordesconto = ""
        txtvalorunitariodesc = txtvalorunitario
End If
ProcAtualizavalores

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

Private Sub txtdesconto2_Change()
On Error GoTo tratar_erro

If Chk_desc2.Value = 1 Then
    If txtdesconto2.Text <> "" Then
        VerifNumero = txtdesconto2.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtdesconto2.Text = ""
            txtdesconto2.SetFocus
            Exit Sub
        End If
        valor = txtdesconto2
        If valor > 100 Then
            USMsgBox ("O desconto não pode ser maior que 100."), vbExclamation, "CAPRIND v5.0"
            txtdesconto2 = ""
            txtdesconto2.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaDesconto2
ElseIf Chk_valor_desc2.Value = 0 Then
        txtdesconto2 = ""
        txtvalordesconto2 = ""
        txtvalorunitariodesc2 = txtvlrunitservico
End If
ProcCalculaValoresServicos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto2_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtdesconto2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto2_LostFocus()
On Error GoTo tratar_erro

If txtdesconto2 = "" Then txtdesconto2 = 0
txtvalordesconto2 = Format(txtvalordesconto2, "###,##0.0000000000")
txtvalorunitariodesc2 = Format(txtvalorunitariodesc2, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEmail_LostFocus()
On Error GoTo tratar_erro

If txtEmail.Text <> "" Then txtEmail.Text = LCase(txtEmail.Text)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspessura_Change()
On Error GoTo tratar_erro

If txtespessura <> "" Then
    VerifNumero = txtespessura
    ProcVerificaNumero
    If VerifNumero = False Then
        txtespessura = ""
        txtespessura.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDcliente_Change()
On Error GoTo tratar_erro

ProcLimpaCliente
If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    ProcPuxaClientes
    'ProcLocalizaVendedorInterno
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtiss_Change()
On Error GoTo tratar_erro

If txtiss.Text <> "" Then
    VerifNumero = txtiss.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtiss.Text = ""
        txtiss.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValoresServicos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtiss_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtiss

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtiss_LostFocus()
On Error GoTo tratar_erro

txtiss.Text = Format(txtiss.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLargura_Change()
On Error GoTo tratar_erro

If txtLargura <> "" Then
    VerifNumero = txtLargura
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLargura = ""
        txtLargura.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtlocal_cobranca_Click()
On Error GoTo tratar_erro

If txtlocal_cobranca <> "" Then Txt_ID_cobranca = txtlocal_cobranca.ItemData(txtlocal_cobranca.ListIndex) Else Txt_ID_cobranca = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtlocal_entrega_Click()
On Error GoTo tratar_erro

If txtlocal_entrega <> "" Then Txt_ID_entrega = txtlocal_entrega.ItemData(txtlocal_entrega.ListIndex) Else Txt_ID_entrega = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNomenclatura_Change()
On Error GoTo tratar_erro

If OPTnovo.Value = 0 And OPTnovoman.Value = 0 Then
    ProcLiberaTabsProd
    ProcLimparProdutos True
    If txtNomenclatura <> "" Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from projproduto where desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            ProcBloqueiaCampos
        Else
            Procliberacampos
        End If
        TBFI.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNomenclatura_LostFocus()
On Error GoTo tratar_erro

If txtNomenclatura <> "" And OPTnovo.Value = 0 And OPTnovoman.Value = 0 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        ProcBloqueiaCampos
    Else
        Procliberacampos
    End If
    TBProduto.Close
Else
    Procliberacampos
End If

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

Private Sub txtNreg1_Change()
On Error GoTo tratar_erro

If txtNreg1 <> "" Then
    VerifNumero = txtNreg1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg1 = ""
        txtNreg1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg2_Change()
On Error GoTo tratar_erro

If txtNreg2 <> "" Then
    VerifNumero = txtNreg2
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg2 = ""
        txtNreg2.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr2_Change()
On Error GoTo tratar_erro

If txtPagIr2 <> "" Then
    VerifNumero = txtPagIr2
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr2 = ""
        txtPagIr2.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr1_Change()
On Error GoTo tratar_erro

If txtPagIr1 <> "" Then
    VerifNumero = txtPagIr1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr1 = ""
        txtPagIr1.SetFocus
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

Private Sub txtPrazo_Produto_Change()
On Error GoTo tratar_erro

If txtPrazo_Produto.Text <> "" Then
    VerifNumero = txtPrazo_Produto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPrazo_Produto.Text = ""
        txtPrazo_Produto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPrazo_Produto_LostFocus()
On Error GoTo tratar_erro

txtPrazo_Produto = Format(txtPrazo_Produto, "###,##0")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPrazo_Servico_Change()
On Error GoTo tratar_erro

If txtPrazo_Servico.Text <> "" Then
    VerifNumero = txtPrazo_Servico.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPrazo_Servico.Text = ""
        txtPrazo_Servico.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPrazo_Servico_LostFocus()
On Error GoTo tratar_erro

txtPrazo_Servico = Format(txtPrazo_Servico, "###,##0")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtservico_Change()
On Error GoTo tratar_erro

TotalServicos = 0
VltUnit = 0
qt = 0
If txtqtservico.Text <> "" Then
    VerifNumero = txtqtservico.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtqtservico.Text = ""
        txtqtservico.SetFocus
        Exit Sub
    End If
    If txtvlrunitservico <> "" Then
        VltUnit = txtvlrunitservico
        qt = txtqtservico
        vlttotal = VltUnit * qt
        txtvlrtotalservico = Format(vlttotal, "###,##0.00")
    End If
End If
ProcCalculaDesconto2
ProcCalculaValoresServicos
TotalServicos = 0
VltUnit = 0
qt = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtservico_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtqtservico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtservico_LostFocus()
On Error GoTo tratar_erro

txtqtservico.Text = Format(txtqtservico.Text, "###,##0.0000")
txtvalordesconto2 = Format(txtvalordesconto2, "###,##0.0000000000")
txtvalorunitariodesc2 = Format(txtvalorunitariodesc2, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtQuantidade.Text <> "" Then
    VerifNumero = txtQuantidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade.Text = ""
        txtQuantidade.SetFocus
        Exit Sub
    End If
End If
ProcCalculaDesconto
ProcAtualizavalores

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
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

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
  
txtTotaldesconto = Format(txtTotaldesconto, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTotaldesconto_LostFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtTotaldesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTransportadora_Change()
On Error GoTo tratar_erro

If txtTransportadora.Text <> "" Then txtidTransportadora = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtuf_Click()
On Error GoTo tratar_erro

If txtuf.Text = "EX" Then
    cmbCidade.Visible = False
    txtCidade.Visible = True
    Cmb_cidade_servico.Clear
Else
    cmbCidade.Visible = True
    txtCidade.Visible = False
    ProcCarregaComboCidade cmbCidade, "Sigla_UF = '" & txtuf & "'", False
    ProcCarregaComboCidade Cmb_cidade_servico, "Sigla_UF = '" & txtuf & "'", True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtunservico_Click()
On Error GoTo tratar_erro

If txtunservico <> "" Then ProcLibera_UN_Com txtunservico, Cmb_un_com_serv

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
If txtvalorunitario <> "0" And txtvalordesconto <> "0" Then
    ProcCalculaValorDesconto
End If

ElseIf Chk_desc.Value = 0 Then
        txtDesconto = ""
        txtvalordesconto = ""
        txtvalorunitariodesc = txtvalorunitario
End If
ProcAtualizavalores

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

Private Sub txtvalordesconto2_Change()
On Error GoTo tratar_erro

If Chk_valor_desc2.Value = 1 Then
    If txtvalordesconto2.Text <> "" Then
        VerifNumero = txtvalordesconto2.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtvalordesconto2.Text = ""
            txtvalordesconto2.SetFocus
            Exit Sub
        End If
        valor = IIf(txtvlrunitservico = "", 0, txtvlrunitservico)
        Valor_Produto = txtvalordesconto2
        If Valor_Produto > valor Then
            USMsgBox ("O valor do desconto não pode ser maior que o valor unitário."), vbExclamation, "CAPRIND v5.0"
            txtvalordesconto2 = ""
            txtvalordesconto2.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaValorDescontoServ
ElseIf Chk_desc2.Value = 0 Then
        txtdesconto2 = ""
        txtvalordesconto2 = ""
        txtvalorunitariodesc2 = txtvlrunitservico
End If
ProcCalculaValoresServicos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto2_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalordesconto2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto2_LostFocus()
On Error GoTo tratar_erro

If txtvalordesconto2 = "" Then txtvalordesconto2 = 0
txtvalordesconto2 = Format(txtvalordesconto2, "###,##0.0000000000")
txtvalorunitariodesc2 = Format(txtvalorunitariodesc2, "###,##0.0000000000")

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
ProcCalculaDesconto
ProcAtualizavalores

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

Private Sub txtVE_LostFocus()
On Error GoTo tratar_erro
 
txtregiao.Text = ""
If txtVE.Text <> "" Then
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select regiao from vendas_vendedores where n_vendedor = " & txtVE.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        txtregiao.Text = TBVendas!regiao
    End If
    TBVendas.Close
End If
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvend_Int_Change()
On Error GoTo tratar_erro

'Set TBClientes = CreateObject("adodb.recordset")
'TBClientes.Open "Select * from Vendas_Vendedores where Vendedor = '" & txtvend_Int & "'", Conexao, adOpenKeyset, adLockOptimistic
''===============================================================
'' Se pertencer ao vendedor
''===============================================================
'If TBClientes.EOF = False Then
'txtComissao.Text = IIf(IsNull(TBClientes!Comissao), 0, TBClientes!Comissao)
'txtVI.Text = TBClientes!ID
'txtvend_Int.Text = TBClientes!vendedor
'TBClientes.Close
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrunitservico_Change()
On Error GoTo tratar_erro

If txtvlrunitservico.Text <> "" Then
    VerifNumero = txtvlrunitservico.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvlrunitservico.Text = ""
        txtvlrunitservico.SetFocus
        Exit Sub
    End If
    If txtqtservico = "" Then Exit Sub
    VltUnit = txtvlrunitservico.Text
    qt = txtqtservico.Text
    vlttotal = VltUnit * qt
End If
ProcCalculaDesconto2
ProcCalculaValoresServicos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrunitservico_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvlrunitservico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrunitservico_LostFocus()
On Error GoTo tratar_erro

txtvlrunitservico = Format(txtvlrunitservico.Text, "###,##0.0000000000")
txtvalordesconto2 = Format(txtvalordesconto2, "###,##0.0000000000")
txtvalorunitariodesc2 = Format(txtvalorunitariodesc2, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdlistaproduto_Click()
On Error GoTo tratar_erro

ProcLiberaTabsProd
PI_Produtos = True
PI_Servicos = False
Vendas_Programacao = False
frmVendas_ListaProduto.Show 1
txtNomenclatura.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_prod()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus <> "ABERTA EM ANALISE" And txtStatus <> "VENDIDA" And txtStatus <> "VENDIDA PARCIAL" And txtStatus <> "FATURADA PARCIAL" Then
    USMsgBox ("Só é permitido criar novo produto em " & IIf(Vendas_PI = True, "pedido interno", "proposta") & " com o status aberta em análise, vendida parcial ou faturada parcial."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, IIf(Vendas_Proposta = True, "proposta", "pedido interno"), "produto", IIf(Vendas_Proposta = True, False, True)) = False Then Exit Sub
Novo_PI1 = True
ProcLimparProdutos False
Frame1(10).Enabled = True
Frame1(12).Enabled = True
SSTab2.Tab = 0
txtNomenclatura.SetFocus
ProcLiberaTabsProd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelaPI()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtCotacao = "" Then
    If Vendas_Proposta = True Then
        Acao = "cancelar o pedido interno"
        NomeCampo = "a proposta comercial"
    Else
        Acao = "cancelar"
        NomeCampo = "o pedido interno"
    End If
    ProcVerificaAcao
    Exit Sub
End If
If USMsgBox("Deseja realmente cancelar o pedido interno " & txtCotacao & "- Rev." & txtrevisao & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If txtStatus.Text <> "VENDIDA" And txtStatus.Text <> "VENDIDA PARCIAL" Then
        USMsgBox ("Só é permitido cancelar pedido interno de proposta com o status vendida ou vendida parcial."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Vendas_PI = True Then
        If FunVerifValidacaoRegistro("cancelar", txtDtValidacao, "mesmo", "o pedido interno", True) = False Then Exit Sub
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select VC.Codigo from vendas_carteira VC INNER JOIN Producao_pedidos PP ON PP.IDcarteira = VC.Codigo where VC.cotacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            USMsgBox ("Não é permitido cancelar este pedido interno, pois já foi emitida ordem de produção para o mesmo."), vbExclamation, "CAPRIND v5.0"
            TBAbrir.Close
            Exit Sub
        End If
        TBAbrir.Close
    End If
    
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_carteira where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        Do While TBVendas.EOF = False
            TBVendas!Liberacao = "ABERTA EM ANALISE"
            TBVendas!Datavendas = Null
            TBVendas!PrazoFinal = Null
            TBVendas!Prazo_original = Null
            TBVendas!PCCliente = Null
            TBVendas.Update
            
            ProcExcluirEmpenhos Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, True
            TBVendas.MoveNext
        Loop
    End If
    Conexao.Execute "UPDATE vendas_proposta Set Status = 'ABERTA EM ANALISE', Datavendas = Null, Tipo = 'PR' where cotacao = " & txtId
    USMsgBox ("Pedido interno " & txtCotacao & "- Rev. " & txtrevisao & " cancelado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Cancelar pedido interno"
    ID_documento = txtId
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    FunAtualizaStatusPropPI txtId
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista2 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista2)
        Lista.SetFocus
    End If
1:
    ProcLimpar
    ProcLimparTudo
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_prod()
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
                If USMsgBox("Deseja realmente excluir este(s) produto(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from vendas_carteira where Codigo = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from vendas_carteira_alteracoes where ID_carteira = " & .ListItems(InitFor) & " and Tipo = '" & IIf(Vendas_PI = True, "VPI", "VPR") & "'"

            '==================================
            Modulo = Formulario
            Evento = "Excluir produto"
            ID_documento = .ListItems(InitFor)
            Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & txtCotacao & " - Rev.: " & txtrevisao
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
            
            'Excluir cliente do produto
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select vendas_proposta.* from (vendas_carteira INNER JOIN vendas_proposta on Vendas_carteira.cotacao = Vendas_proposta.cotacao) INNER JOIN projproduto on vendas_carteira.desenho = projproduto.desenho where Vendas_carteira.desenho = '" & .ListItems(InitFor).SubItems(1) & "' and Vendas_proposta.IDcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = True Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from projproduto where desenho = '" & .ListItems(InitFor).SubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select tbl_dados_nota_fiscal.*, tbl_detalhes_nota.codproduto  from tbl_detalhes_nota INNER JOIN tbl_dados_nota_fiscal on tbl_detalhes_nota.id_nota = tbl_dados_nota_fiscal.ID where tbl_detalhes_nota.int_Cod_Produto = '" & .ListItems(InitFor).SubItems(1) & "' and tbl_dados_nota_fiscal.Id_Int_Cliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then Conexao.Execute "DELETE from Projproduto_clientes WHERE codproduto = " & TBItem!Codproduto & " and IDcliente = " & txtIDcliente
                    TBFI.Close
                End If
                TBItem.Close
            End If
            TBOrdem.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimparProdutos False
    Listprod.ListItems.Clear
    ProcAtualizalistaProdutos (1)
    Listprod.SetFocus
    Frame1(10).Enabled = False
    Frame1(12).Enabled = False
    SSTab2.Tab = 0
    Novo_PI1 = False
    
    If Vendas_PI = True Then
        If FunAtualizaStatusPropPI(txtId) = True Then
            SSTab1.Tab = 0
            SSTab1_Click (0)
            ProcLimpar
        End If
    Else
        FunAtualizaStatusPropPI txtId
    End If
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValores()
On Error GoTo tratar_erro

If txtQuantidade.Text = "" Or txtvalorunitario.Text = "" Then
    txtdbl_valoripi = "0,00"
    txtvalor_icms = "0,00"
    Exit Sub
End If
'Zera valores
SumICMS = 0
SumIPI = 0
SumTotNota = 0
SumTotProdutos = 0
VlrIPI = 0

'Se tem IPI
If TemIPI = "SIM" Then
    txtInt_ipi.Text = IntIPI
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Clientes_Impostos where IDCliente = " & txtIDcliente & " and ID_CF = " & IIf(Txt_ID_CF = "", 0, Txt_ID_CF), Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        If FunVerifCalcIPISDesc(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            VlrIPI = txtvalorunitario
            If TBFIltro!PorcentagemIPI <> 0 Then VlrIPI = VlrIPI / TBFIltro!PorcentagemIPI
            VlrIPI = (VlrIPI - txtvalorunitario.Text) * txtQuantidade
        Else
            VlrIPI = txtvalorunitariodesc
            If TBFIltro!PorcentagemIPI <> 0 Then VlrIPI = VlrIPI / TBFIltro!PorcentagemIPI
            VlrIPI = (VlrIPI - txtvalorunitariodesc.Text) * txtQuantidade
        End If
    Else
        If FunVerifCalcIPISDesc(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            VlrIPI = txtvalorunitario * txtQuantidade
        Else
            VlrIPI = txtvalorunitariodesc * txtQuantidade
        End If
        VlrIPI = Format((VlrIPI * IntIPI) / 100, "###,##0.00")
    End If
    TBFIltro.Close
    txtdbl_valoripi = Format(VlrIPI, "###,##0.00")
Else
    txtInt_ipi.Text = 0
    txtdbl_valoripi = "0,00"
End If

'Atribui valores
'===========================================================
'If FRETE_ICMS = True Then
'txtvalor_total = Format(txtvalorunitariodesc.Text * txtQuantidade + IIf(txtvFrete <> "", txtvFrete, "0"), "###,##0.00")
'Else
txtvalor_total = Format(txtvalorunitariodesc.Text * txtQuantidade, "###,##0.00")
'End If
'===========================================================

'Se tem icms
If TemICMS = "SIM" Then
    txtint_icms.Text = IntICMS
    SumTotProdutos = SumTotProdutos + Format(txtvalor_total.Text, "###,##0.00")
Else
    txtint_icms.Text = 0
End If

'=============================================================================
valor = IIf(txtvalor_total = "", 0, txtvalor_total.Text)
Valor1 = IIf(txtvFrete = "", 0, txtvFrete.Text)
Valor2 = valor + Valor1
ProcCalculaBC Cmb_empresa.ItemData(Cmb_empresa.ListIndex), IIf(Txt_CFOP_prod = "", "0.000", Txt_CFOP_prod), 0, Valor2, txtdbl_valoripi, SomarIPI, SomarIPIST, TemReducaoBC, False, Cmb_CST_ICMS, "P", 0, ""
valor = 0
Valor1 = 0
Valor2 = 0
'============================================================================
txtvalor_icms = Format((BC * txtint_icms.Text) / 100, "###,##0.00")
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValoresServicos()
On Error GoTo tratar_erro

VlISS = 0
TotalServicos = 0
VltUnit = 0
qt = 0
If txtqtservico.Text = "" Or txtvlrunitservico.Text = "" Then
    txtvlrtotalservico.Text = "0,00"
    Exit Sub
End If
'Atribui valores
txtvlrtotalservico = Format(txtvalorunitariodesc2.Text * txtqtservico, "###,##0.00")
VltUnit = txtvalorunitariodesc2.Text
If txtiss.Text <> "" Then
    VlISS = txtiss.Text
    qt = txtqtservico.Text
    TotalServicos = VltUnit * qt
    txtvlrISS.Text = Format((TotalServicos * VlISS) / 100, "###,##0.00")
    txtvlrtotalservico = Format(TotalServicos, "###,##0.00")
Else
    txtvlrISS.Text = "0,00"
End If
VlISS = 0
TotalServicos = 0
VltUnit = 0
qt = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProduto()
On Error GoTo tratar_erro

If txttipocliente = "JP" Or txttipocliente = "FP" Then
    txtNomenclatura = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtNomenclatura, txtreferencia, IIf(txtRev_cod = "", 0, txtRev_cod), txtdesctecnica, txtEspecificacoes, cmbfamilia, IIf(txtvalorunitario = "", 0, txtvalorunitario), 0, 0, cmbun, Cmb_un_com, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), False, True, True, False, 1, "P", Txt_observacoes_prod, IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, txtIDcliente, txtCliente, "C")
Else
    txtNomenclatura = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtNomenclatura, txtreferencia, IIf(txtRev_cod = "", 0, txtRev_cod), txtdesctecnica, txtEspecificacoes, cmbfamilia, 0, IIf(txtvalorunitario = "", 0, txtvalorunitario), 0, cmbun, Cmb_un_com, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), False, True, True, False, 1, "P", Txt_observacoes_prod, IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, txtIDcliente, txtCliente, "C")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procnovoservico()
On Error GoTo tratar_erro

If txttipocliente = "JP" Or txttipocliente = "FP" Then
    txtcodservico = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtcodservico, txtReferencia_serv, IIf(txtRev_serv = "", 0, txtRev_serv), txtdescservico, txtdesccomservico, cmbfamiliaservico, IIf(txtvlrunitservico = "", 0, txtvlrunitservico), 0, 0, txtunservico, Cmb_un_com_serv, 0, False, True, True, False, 5, "S", txtObs_serv, 0, 0, 0, "", txtIDcliente, txtCliente, "C")
Else
    txtcodservico = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtcodservico, txtReferencia_serv, IIf(txtRev_serv = "", 0, txtRev_serv), txtdescservico, txtdesccomservico, cmbfamiliaservico, 0, IIf(txtvlrunitservico = "", 0, txtvlrunitservico), 0, txtunservico, Cmb_un_com_serv, 0, False, True, True, False, 5, "S", txtObs_serv, 0, 0, 0, "", txtIDcliente, txtCliente, "C")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoMan()
On Error GoTo tratar_erro

If txttipocliente = "JP" Or txttipocliente = "FP" Then
    txtNomenclatura = FunCriaNovoProdServ(True, "", txtNomenclatura, txtreferencia, IIf(txtRev_cod = "", 0, txtRev_cod), txtdesctecnica, txtEspecificacoes, cmbfamilia, txtvalorunitario, 0, 0, cmbun, Cmb_un_com, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), False, True, True, False, 1, "P", Txt_observacoes_prod, IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, txtIDcliente, txtCliente, "C")
Else
    txtNomenclatura = FunCriaNovoProdServ(True, "", txtNomenclatura, txtreferencia, IIf(txtRev_cod = "", 0, txtRev_cod), txtdesctecnica, txtEspecificacoes, cmbfamilia, 0, txtvalorunitario, 0, cmbun, Cmb_un_com, IIf(Txt_ID_CF = "", 0, Txt_ID_CF), False, True, True, False, 1, "P", Txt_observacoes_prod, IIf(txtComprimento = "", 0, txtComprimento), IIf(txtLargura = "", 0, txtLargura), IIf(txtespessura = "", 0, txtespessura), txtDureza, txtIDcliente, txtCliente, "C")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoServicoMan()
On Error GoTo tratar_erro

If txttipocliente = "JP" Or txttipocliente = "FP" Then
    txtcodservico = FunCriaNovoProdServ(True, "", txtcodservico, txtReferencia_serv, IIf(txtRev_serv = "", 0, txtRev_serv), txtdescservico, txtdesccomservico, cmbfamiliaservico, txtvlrunitservico, 0, 0, txtunservico, Cmb_un_com_serv, 0, False, True, True, False, 5, "S", txtObs_serv, 0, 0, 0, "", txtIDcliente, txtCliente, "C")
Else
    txtcodservico = FunCriaNovoProdServ(True, "", txtcodservico, txtReferencia_serv, IIf(txtRev_serv = "", 0, txtRev_serv), txtdescservico, txtdesccomservico, cmbfamiliaservico, 0, txtvlrunitservico, 0, txtunservico, Cmb_un_com_serv, 0, False, True, True, False, 5, "S", txtObs_serv, 0, 0, 0, "", txtIDcliente, txtCliente, "C")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_prod()
On Error GoTo tratar_erro

If txtDtValidacao.Text <> "" And LiberarAlteracao = False Then
    USMsgBox ("Atenção não é permitido alterar pedido validado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(10).Enabled = False Or Frame1(12).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If OPTnovo.Value = 0 And OPTnovoman.Value = 0 And txtNomenclatura = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtNomenclatura.SetFocus
    Exit Sub
End If
If txtdesctecnica.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdesctecnica.SetFocus
    Exit Sub
End If

If Vendas_PI = True Then
    With txtpccliente
        If .Text = "" Then
            NomeCampo = "o pedido do cliente"
            ProcVerificaAcao
            .Locked = False
            .TabStop = True
            .SetFocus
            Exit Sub
        End If
    End With
    If IsDate(mskprazo) = False Then
        NomeCampo = "o prazo final"
        ProcVerificaAcao
        mskprazo.SetFocus
        Exit Sub
    End If
Else
    If txtPrazo_Produto = "" Then
        NomeCampo = "o prazo em dias"
        ProcVerificaAcao
        txtPrazo_Produto.SetFocus
        Exit Sub
    Else
        Valor_Cofins_Prod = txtPrazo_Produto
        If Valor_Cofins_Prod - Int(Valor_Cofins_Prod) > 0 Then
            USMsgBox ("Só é permitido número inteiro no prazo em dias."), vbExclamation, "CAPRIND v5.0"
            txtPrazo_Produto.SetFocus
            Exit Sub
        End If
    End If
End If

If Txt_data_retorno <> "__/__/____" And IsDate(Txt_data_retorno) = False Then
    NomeCampo = "a data prevista do retorno"
    ProcVerificaAcao
    Txt_data_retorno.SetFocus
    Exit Sub
End If

If txtEspecificacoes.Text = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtEspecificacoes.SetFocus
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
If cmbfamilia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
valor = IIf(txtQuantidade = "", 0, txtQuantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQuantidade.SetFocus
    Exit Sub
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

If Vendas_PI = True Then
    If OPTnovo.Value = 0 And OPTnovoman.Value = 0 Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where desenho = '" & txtNomenclatura.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            USMsgBox ("Não foi encontrado nenhum produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
            txtNomenclatura.SetFocus
            TBProduto.Close
            Exit Sub
        End If
        TBProduto.Close
    End If
End If

If Txt_ID_CFOP_prod <> "" And Txt_ID_CFOP_prod <> "0" And txtuf <> "" And txtuf <> "EX" Then
    If FunVerificaCFOPUF(Txt_ID_CFOP_prod, txtuf) = False Then Exit Sub
End If

If Txt_ID_CFOP_prod <> "" And Txt_ID_CFOP_prod <> "0" And txtuf <> "" And txtuf <> "EX" Then
        If Cmb_CST_ICMS.Text = "" Then
        NomeCampo = "CST do ICMS"
        ProcVerificaAcao
        Cmb_CST_ICMS.SetFocus
        Exit Sub
    End If
End If

If Txt_CF.Text = "" Then
        NomeCampo = "NCM"
        ProcVerificaAcao
        Txt_CF.SetFocus
        Exit Sub
End If

'Verifica se ja existe o mesmo produto na proposta
If Novo_PI1 = True And txtNomenclatura <> "" And OPTnovo.Value = 0 And OPTnovoman.Value = 0 Then
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from vendas_carteira where cotacao = " & txtId & " and Desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        USMsgBox ("Já existe um produto com o código " & txtNomenclatura & IIf(Vendas_Proposta = True, " nessa proposta", " nesse pedido") & "."), vbExclamation, "CAPRIND v5.0"
        If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            TBCotacao.Close
            Exit Sub
        End If
    End If
End If

'Se for novo produto
If OPTnovo.Value = 1 Then
    Call ProcNovoProduto
    If txtreferencia <> "" Then
        cmbReferencia.AddItem txtreferencia
        cmbReferencia = txtreferencia
    End If
    OPTnovo.Value = 0
End If
If OPTnovoman.Value = 1 Then
    If txtNomenclatura.Text = "" Then
        USMsgBox ("Informe o código interno antes de salvar."), vbExclamation, "CAPRIND v5.0"
        txtNomenclatura.SetFocus
        Exit Sub
    End If
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtNomenclatura.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtNomenclatura.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    Call ProcNovoProdutoMan
    If txtreferencia <> "" Then
        cmbReferencia.AddItem txtreferencia
        cmbReferencia = txtreferencia
    End If
    OPTnovoman.Value = 0
End If

If OPTnovo.Value = 0 And OPTnovoman.Value = 0 Then Conexao.Execute "Update projproduto Set RevDesenho = '" & IIf(txtRev_cod.Text = "", 0, txtRev_cod.Text) & "' where desenho = '" & txtNomenclatura.Text & "'"

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * from vendas_carteira where Codigo = " & txtid_produto, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    If Vendas_PI = True Then
        If TBCotacao!Liberacao <> "VENDIDA" And TBCotacao!Liberacao <> "VENDIDA PARCIAL" And TBCotacao!Liberacao <> "FATURADO PARCIAL" Then
            USMsgBox ("Só é permitido alterar produto com o status vendido, vendido parcial ou faturado parcial."), vbExclamation, "CAPRIND v5.0"
            TBCotacao.Close
            Exit Sub
        End If
        
'===========================================================
' Se for trocar o produto
'===========================================================
        If TBCotacao!Desenho <> txtNomenclatura Then
            If FunVerifAltCodQtde(True, True) = False Then Exit Sub
        End If
'===========================================================
' Se for trocar a quantidade
'===========================================================
        valor = txtQuantidade
        If TBCotacao!quantidade <> valor Then
            If FunVerifAltCodQtde(True, False) = False Then Exit Sub
        End If
    Else
        If TBCotacao!Liberacao <> "ABERTA EM ANALISE" Then
            USMsgBox ("Só é permitido alterar produto com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
            TBCotacao.Close
            Exit Sub
        End If
    End If
Else
    
     'Busca o valor cadastrado de limite de credito no cliente
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select txtLimiteCredito from Clientes where idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then LimiteCredito = TBClientes!txtLimiteCredito
    
    TBClientes.Close
    
    If LimiteCredito <> 0 Then
        'Totaliza o limite utilizado pra contas em aberto
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Sum(valor) as LimiteCreditoUtilizadoAberto, idCliente from tbl_contas_receber where idCliente = " & txtIDcliente & " AND Status = 'TÍTULO EM ABERTO' GROUP BY idCliente", Conexao, adOpenKeyset, adLockReadOnly
        If TBContas.EOF = False Then
            LimiteCreditoUtilizadoAberto = TBContas!LimiteCreditoUtilizadoAberto
        Else
            LimiteCreditoUtilizadoAberto = 0
        End If
        TBContas.Close
        
        'Totaliza o limite utilizado pra contas parcial
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Sum(valorpendente) as LimiteCreditoUtilizadoParcial, idCliente from tbl_contas_receber where idCliente = " & txtIDcliente & " AND Status = 'TÍTULO RECEBIDO PARCIAL' and LogSit = 'N' GROUP BY idCliente", Conexao, adOpenKeyset, adLockReadOnly
        If TBContas.EOF = False Then
            LimiteCreditoUtilizadoParcial = TBContas!LimiteCreditoUtilizadoParcial
        Else
            LimiteCreditoUtilizadoParcial = 0
        End If
        TBContas.Close
        
        'Totalliza o saldo do limite utilizado
        
        LimiteCreditoSaldo = LimiteCreditoUtilizadoAberto + LimiteCreditoUtilizadoParcial + txtvalor_total
        LimiteDisponivel = LimiteCredito - (LimiteCreditoUtilizadoAberto + LimiteCreditoUtilizadoParcial)
        
        If LimiteCredito <= LimiteCreditoSaldo Then
            MsgBox ("O limite de credito desse cliente de R$" & LimiteCredito & " foi atingido, o limite disponível é de R$" & LimiteDisponivel & ", não é possivel adicionar esse produto!"), vbCritical + vbOKOnly
            Exit Sub
        End If
    End If
'==========================================
' Cadastra um novo produto
'==========================================
    TBCotacao.AddNew
End If

TBCotacao!Tem_ordem = False
TBCotacao!Liberacao = "VENDIDA"
TBCotacao!Datavendas = IIf(IsDate(txtDatavendas_PI) = True, txtDatavendas_PI, Now)
TBCotacao!Observacoes = Txt_observacoes_prod
TBCotacao!Obs_faturamento = Txt_observacoes_fat_prod
TBCotacao!IDAnalise = IDAnalise
TBCotacao!Desenho = txtNomenclatura.Text
TBCotacao!N_referencia = cmbReferencia
TBCotacao!ID_CFOP = IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod)
TBCotacao!quantidade = txtQuantidade.Text
TBCotacao!Qtde_produzir = TBCotacao!quantidade / FunVerificaTabelaConversaoUnidade(cmbun, Cmb_un_com)
TBCotacao!Rev_codinterno = IIf(txtRev_cod.Text = "", 0, txtRev_cod)
TBCotacao!Descricao = Trim(txtEspecificacoes.Text)

If chkNovo_projeto.Value = 1 Then
    TBCotacao!Novo_projeto = True
Else
    TBCotacao!Novo_projeto = False
End If

If Chk_utiliza_mat_consignado.Value = 1 Then
    TBCotacao!Utiliza_mat_cons = True
Else
    TBCotacao!Utiliza_mat_cons = False
End If

TBCotacao!Comprimento = IIf(txtComprimento = "", Null, txtComprimento)
TBCotacao!Largura = IIf(txtLargura = "", Null, txtLargura)
TBCotacao!Espessura = IIf(txtespessura = "", Null, txtespessura)
TBCotacao!Dureza = txtDureza
TBCotacao!descricao_tecnica = Trim(txtdesctecnica.Text)
TBCotacao!Unidade = cmbun.Text
TBCotacao!Unidade_com = Cmb_un_com.Text
TBCotacao!Familia = cmbfamilia.Text
TBCotacao!ID_CF = IIf(Txt_ID_CF = "", Null, Txt_ID_CF)
TBCotacao!txt_CST = Cmb_CST_ICMS
TBCotacao!Cotacao = txtId.Text
TBCotacao!Tipo = "P"
TBCotacao!Prioridade = Cmb_prioridade

If chkRetorno.Value = 1 Then
    TBCotacao!retorno = True
Else
    TBCotacao!retorno = False
End If

TBCotacao!Data_retorno = IIf(Txt_data_retorno = "__/__/____", Null, Txt_data_retorno)
TBCotacao!Inspecao = IIf(txtinspecao = "", Null, txtinspecao)
TBCotacao!Embalagem = IIf(txtembalagem = "", Null, txtembalagem)
TBCotacao!Gravacao = IIf(txtGravacao = "", Null, txtGravacao)
TBCotacao!preco_unitario = IIf(txtvalorunitario = "", 0, txtvalorunitario)
TBCotacao!Desconto = IIf(txtDesconto = "", 0, txtDesconto)
TBCotacao!ValorDesconto = IIf(txtvalordesconto = "", 0, txtvalordesconto)
TBCotacao!preco_unitario_desconto = txtvalorunitariodesc
TBCotacao!vFrete = IIf(txtvFrete = "", 0, txtvFrete)
TBCotacao!preco_lote = txtvalor_total.Text
TBCotacao!N_item = Trim(Txt_n_item)

'Calcula comissão
Qtde = 0
Qtd = 0
If txtComissao <> "" Then
    Qtde = txtComissao
    Qtd = TBCotacao!preco_lote
    Qtd = (Qtd * Qtde) / 100
    
    TBCotacao!Comissao = txtComissao
    TBCotacao!ValorComissao = Qtd
Else
    TBCotacao!Comissao = 0
    TBCotacao!ValorComissao = 0
End If
'==================================
txtvalor_total.Text = Format(TBCotacao!preco_lote, "###,##0.00")
'==============================================
TBCotacao!dbl_valoripi = Format(txtdbl_valoripi.Text, "###,##0.00")
TBCotacao!IntICMS = txtint_icms
TBCotacao!int_IPI = txtInt_ipi
TBCotacao!dbl_Valor_ICMS = Format(txtvalor_icms.Text, "###,##0.00")

TBCotacao!BC_ICMS = 0
TBCotacao!BC_ICMS_ST = 0
TBCotacao!Valor_ICMS_ST = 0
If Txt_ID_CF <> "" Then
    ProcValorImposto txtCotacao, Txt_ID_CF, IIf(txtIDcliente = "", 0, txtIDcliente), txtCliente, txtuf, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), RegimeEmpresa_PI
    ProcControleImposto IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), IIf(txtIDcliente = "", 0, txtIDcliente)
'=================================================================
valor = IIf(txtvalor_total = "", 0, txtvalor_total.Text)
Valor1 = IIf(txtvFrete = "", 0, txtvFrete.Text)
Valor2 = valor + Valor1
ProcCalculaBC Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Txt_CFOP_prod, 0, Valor2, txtdbl_valoripi, SomarIPI, SomarIPIST, TemReducaoBC, False, Cmb_CST_ICMS, "P", 0, ""
If TemICMS = "SIM" And TBCotacao!dbl_Valor_ICMS <> 0 Then TBCotacao!BC_ICMS = BC
valor = 0
Valor1 = 0
Valor2 = 0
'==================================================================================
    If Cmb_CST_ICMS <> "" And chkRetorno = 0 Then
        ProcSubstituicaoTributaria txtuf, Cmb_CST_ICMS, Txt_ID_CF, IIf(txtIDcliente = "", 0, txtIDcliente), txtCliente, txtvalorunitariodesc, txtQuantidade, BC, BCST, 0, 0, 0, False, False, 0
        TBCotacao!Valor_ICMS_ST = ICMSCST
        If ICMSCST <> 0 Then TBCotacao!BC_ICMS_ST = BCICMSCST
    End If
End If

'Impostos
Valor_total = txtvalorunitariodesc * txtQuantidade
Valor_IPI = txtdbl_valoripi
If IsNull(TBCotacao!ID_CF) = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_classificacaofiscal where Idclass = " & TBCotacao!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        'Verifica se a CF tem retenção de PIS/Cofins, destaca PIS/Cofins e grava no produto
        If DestacaImpostos = "SIM" And TBFI!Retem_PIS_Cofins = True And chkRetorno.Value = 0 Then
            TBCotacao!Valor_Retencao_PIS = Format((Valor_total * IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)) / 100, "###,##0.00")
            TBCotacao!Valor_Retencao_Cofins = Format((Valor_total * IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)) / 100, "###,##0.00")
        Else
            TBCotacao!Valor_Retencao_PIS = 0
            TBCotacao!Valor_Retencao_Cofins = 0
        End If
        
        'Recalcula o valor do PIS e Cofins de acordo com a alíquota informada na CF se tiver diferente de zero
        If DestacaImpostos = "SIM" Then
            PIS_Prod = IIf(IsNull(TBFI!PIS_destaca), 0, TBFI!PIS_destaca)
            Cofins_Prod = IIf(IsNull(TBFI!Cofins_destaca), 0, TBFI!Cofins_destaca)
            If Regime <> 1 And chkRetorno.Value = 0 Then
                If TemPIS = True And PIS_Prod <> 0 Then
                    TBCotacao!PIS_Prod = PIS_Prod
                    TBCotacao!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00")
                End If
                If TemCOFINS = True And Cofins_Prod <> 0 Then
                    TBCotacao!Cofins_Prod = Cofins_Prod
                    TBCotacao!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00")
                End If
            End If
        Else
            TBCotacao!PIS_Prod = 0
            TBCotacao!Total_PIS_prod = 0
            TBCotacao!Cofins_Prod = 0
            TBCotacao!Total_Cofins_prod = 0
        End If
    End If
    TBFI.Close
End If
TBCotacao!N_Serie = Txt_n_serie

If Vendas_PI = True Then
    TBCotacao!PCCliente = Trim(txtpccliente.Text)
    'Cadastrar automaticamente a alteração da data
    If Novo_PI1 = False And IsNull(TBCotacao!PrazoFinal) = False Then
        If mskprazo <> TBCotacao!PrazoFinal Then ProcINSERTINTO "vendas_carteira_alteracoes", "ID_carteira, Data, Responsavel, Data_alteracao, Responsavel_alteracao, Obs, Alteracao_prazo, Padrao, Tipo", "" & txtid_produto & ",'" & Date & "','" & pubUsuario & "','" & Date & "','" & pubUsuario & "', NULL,'ALTERADO O PRAZO DE ENTREGA DE " & Format(TBCotacao!PrazoFinal, "dd/mm/yy") & " PARA " & Format(mskprazo, "dd/mm/yy") & "', 'True', 'VPI'"
    End If
    If Novo_PI1 = True And IsNull(TBCotacao!Prazo_original) = True Then TBCotacao!Prazo_original = mskprazo.Text
    TBCotacao!PrazoFinal = mskprazo.Text
Else
    'Cadastrar automaticamente a alteração da data
    If Novo_PI1 = False And IsNull(TBCotacao!prazofinaldias) = False Then
        If txtPrazo_Produto <> TBCotacao!prazofinaldias Then ProcINSERTINTO "vendas_carteira_alteracoes", "ID_carteira, Data, Responsavel, Data_alteracao, Responsavel_alteracao, Obs, Alteracao_prazo, Padrao, Tipo", "" & txtid_produto & ",'" & Date & "','" & pubUsuario & "','" & Date & "','" & pubUsuario & "', NULL,'ALTERADO O PRAZO DE ENTREGA DE " & TBCotacao!prazofinaldias & " DIAS PARA " & txtPrazo_Produto & " DIAS', 'True', 'VPR'"
    End If
    TBCotacao!prazofinaldias = IIf(txtPrazo_Produto = "", Null, txtPrazo_Produto)
End If

TBCotacao!Caminho_PCCliente = Caminho_PC_prod_PI
If Chk_antecipacao.Value = 1 Then TBCotacao!Antecipacao_fat = True Else TBCotacao!Antecipacao_fat = False
If Chk_faturamento_parcial.Value = 1 Then TBCotacao!Faturamento_parcial = True Else TBCotacao!Faturamento_parcial = False

'Empresa
If Txt_ID_CFOP_prod <> "" Then ProcControleImposto IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), IIf(txtIDcliente = "", 0, txtIDcliente)
ProcVerifImpostosEmpresa Cmb_empresa.ItemData(Cmb_empresa.ListIndex), chkRetorno, "", False, 0, False, TabelaSN_PI, 0
'Novo cálculo simples nacional 2018
TBCotacao!DAS = DAS
If DAS <> 0 Then TBCotacao!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBCotacao!Total_DAS = 0
TBCotacao!PIS_Prod = PIS_Prod
If PIS_Prod <> 0 Then TBCotacao!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00") Else TBCotacao!Total_PIS_prod = 0
TBCotacao!Cofins_Prod = Cofins_Prod
If Cofins_Prod <> 0 Then TBCotacao!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00") Else TBCotacao!Total_Cofins_prod = 0
TBCotacao!CSLL_Prod = CSLL_Prod
If CSLL_Prod <> 0 Then TBCotacao!Total_CSLL_prod = Format((Valor_total * CSLL_Prod) / 100, "###,##0.00") Else TBCotacao!Total_CSLL_prod = 0
TBCotacao!IRPJ_Prod = IRPJ_Prod
If IRPJ_Prod <> 0 Then TBCotacao!Total_IRPJ_prod = Format((Valor_total * IRPJ_Prod) / 100, "###,##0.00") Else TBCotacao!Total_IRPJ_prod = 0
TBCotacao!cpp = CPP_Prod
If CPP_Prod <> 0 Then TBCotacao!Total_CPP = Format((Valor_total * CPP_Prod) / 100, "###,##0.00") Else TBCotacao!Total_CPP = 0

TBCotacao.Update
txtid_produto = TBCotacao!CODIGO
TBCotacao.Close

If txtIDcliente <> "0" Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select codproduto from Projproduto where desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then ProcAgregarProdutoCli TBFI!Codproduto, txtIDcliente, txttipocliente, cmbun, Cmb_un_com, IIf(txtvalorunitario = "", 0, txtvalorunitario)
    TBFI.Close
End If

'Atualiza valor de venda do produto
'If txtTipoCliente <> "JR" And txtTipoCliente <> "FR" Then
'    ProcAtualizaValorProdServ False, 0, True, txtvalorunitario, 0, txtNomenclatura
'Else
'    ProcAtualizaValorProdServ False, 0, False, 0, txtvalorunitario, txtNomenclatura
'End If

Valor_total = 0
Valor_IPI = 0
ProcAtualizalistaProdutos (IIf(ReturnNumbersOnly(Left(lblPaginas1.Caption, Len(lblPaginas1.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas1.Caption, Len(lblPaginas1.Caption) - 5))))
If Novo_PI1 = True Then
    USMsgBox ("Novo produto cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto"
Else
    Evento = "Alterar produto"
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    If CodigoLista <> 0 And Listprod.ListItems.Count <> 0 Then
        Listprod.SelectedItem = Listprod.ListItems(CodigoLista)
1:
        Listprod.SetFocus
    End If
End If
'==================================
Modulo = Formulario
ID_documento = txtid_produto
Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & txtCotacao & " - Rev.: " & txtrevisao
Documento1 = "Cód. interno: " & txtNomenclatura
ProcGravaEvento
'==================================
FunAtualizaStatusPropPI txtId
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
Novo_PI1 = False
'LiberarAlteracao = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvFrete_Change()
On Error GoTo tratar_erro

If txtvFrete.Text <> "" Then
    VerifNumero = txtvFrete.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvFrete.Text = ""
        txtvFrete.SetFocus
        Exit Sub
    End If
End If
ProcCalculaDesconto
ProcAtualizavalores

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvFrete_LostFocus()
On Error GoTo tratar_erro

txtvFrete = Format(txtvFrete, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarTotais(ID As Long)
On Error GoTo tratar_erro
Dim Valor_total_Frete As Double

'PRODUTOS
TotalProduto = 0
valor = 0
SumIPI = 0
TotalICMSCST = 0
BASECALCULO = 0
TotalICMS = 0
TotalBCICMSCST = 0
Valor_PIS_Prod = 0
Valor_Cofins_Prod = 0
Valor_CSLL_Prod = 0
Valor_IRPJ_Prod = 0
Valor_DAS = 0
Valor_Retencao_PIS = 0
Valor_Retencao_Cofins = 0
Valor_total_Frete = 0

If Vendas_PI = True Then TextoFiltro = " and (Left(Liberacao, 7) = 'VENDIDA' or Left(Liberacao, 7) = 'FATURAR' or Left(Liberacao, 8) = 'FATURADO')" Else TextoFiltro = ""
Set TBTotaisnota = CreateObject("adodb.recordset")
StrSql = "Select Sum(ROUND(preco_unitario * Quantidade, 2)) as TotalProduto,Sum(vFrete) as TotalFrete, sum(Valordesconto) as ValorDesconto, Sum(ROUND(Preco_lote, 2)) as Valor, Sum(dbl_valoripi) as SumIPI, Sum(Valor_ICMS_ST) as TotalICMSCST, Sum(BC_ICMS) as BASECALCULO, Sum(dbl_Valor_ICMS) as TotalICMS, Sum(BC_ICMS_ST) as TotalBCICMSCST, Sum(Total_PIS_prod) as Valor_PIS_Prod, Sum(Total_Cofins_prod) as Valor_Cofins_Prod, Sum(Total_CSLL_prod) as Valor_CSLL_Prod, Sum(Total_IRPJ_prod) as Valor_IRPJ_Prod, Sum(Total_DAS) as Valor_DAS, Sum(Valor_Retencao_PIS) as Valor_Retencao_PIS, Sum(Valor_Retencao_Cofins) as Valor_Retencao_Cofins from vendas_carteira where cotacao = " & ID & " and Tipo = 'P' and Retorno = 'False' " & TextoFiltro
'Debug.print StrSql

TBTotaisnota.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    TotalProduto = IIf(IsNull(TBTotaisnota!TotalProduto), 0, TBTotaisnota!TotalProduto)
    valor = IIf(IsNull(TBTotaisnota!valor), 0, TBTotaisnota!valor) '- IIf(IsNull(TBTotaisnota!TotalFrete), 0, TBTotaisnota!TotalFrete)
'========================================================================
Valor_total_Frete = IIf(IsNull(TBTotaisnota!TotalFrete), 0, TBTotaisnota!TotalFrete)
'========================================================================
    SumIPI = IIf(IsNull(TBTotaisnota!SumIPI), 0, TBTotaisnota!SumIPI)
    TotalICMSCST = IIf(IsNull(TBTotaisnota!TotalICMSCST), 0, TBTotaisnota!TotalICMSCST)
    BASECALCULO = IIf(IsNull(TBTotaisnota!BASECALCULO), 0, TBTotaisnota!BASECALCULO)
    TotalICMS = IIf(IsNull(TBTotaisnota!TotalICMS), 0, TBTotaisnota!TotalICMS)
    TotalBCICMSCST = IIf(IsNull(TBTotaisnota!TotalBCICMSCST), 0, TBTotaisnota!TotalBCICMSCST)
    Valor_PIS_Prod = IIf(IsNull(TBTotaisnota!Valor_PIS_Prod), 0, TBTotaisnota!Valor_PIS_Prod)
    Valor_Cofins_Prod = IIf(IsNull(TBTotaisnota!Valor_Cofins_Prod), 0, TBTotaisnota!Valor_Cofins_Prod)
    Valor_CSLL_Prod = IIf(IsNull(TBTotaisnota!Valor_CSLL_Prod), 0, TBTotaisnota!Valor_CSLL_Prod)
    Valor_IRPJ_Prod = IIf(IsNull(TBTotaisnota!Valor_IRPJ_Prod), 0, TBTotaisnota!Valor_IRPJ_Prod)
    Valor_DAS = IIf(IsNull(TBTotaisnota!Valor_DAS), 0, TBTotaisnota!Valor_DAS)
    Valor_Retencao_PIS = IIf(IsNull(TBTotaisnota!Valor_Retencao_PIS), 0, TBTotaisnota!Valor_Retencao_PIS)
    Valor_Retencao_Cofins = IIf(IsNull(TBTotaisnota!Valor_Retencao_Cofins), 0, TBTotaisnota!Valor_Retencao_Cofins)
    Valor_desconto = IIf(IsNull(TBTotaisnota!ValorDesconto), 0, TBTotaisnota!ValorDesconto)
End If

'Retorno
VlrTotalRetorno = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select Sum(ROUND(preco_unitario * Quantidade, 2)) as VlrTotalRetorno from vendas_carteira where cotacao = " & ID & " and Tipo = 'P' and Retorno = 'True' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    VlrTotalRetorno = IIf(IsNull(TBTotaisnota!VlrTotalRetorno), 0, TBTotaisnota!VlrTotalRetorno)
End If

'SERVIÇOS
TotalServicos = 0
Valor1 = 0
Valor_PIS_Serv = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Serv = 0
TotalISS = 0
Valor_INSS_Serv = 0
Valor_IRPJ_Serv = 0
Valor_IRRF_Serv = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select Sum(ROUND(preco_unitario * Quantidade, 2)) as TotalServicos, Sum(ROUND(Preco_lote, 2)) as Valor1, Sum(Total_PIS_serv) as Valor_PIS_Serv, Sum(Total_Cofins_serv) as Valor_Cofins_Serv, Sum(Total_CSLL_serv) as Valor_CSLL_Serv, Sum(vlriss) as TotalISS, Sum(Total_INSS_serv) as Valor_INSS_Serv, Sum(Total_IRPJ_serv) as Valor_IRPJ_Serv, Sum(Total_DAS) as Valor_DAS, Sum(Total_IRRF_serv) as Valor_IRRF_Serv from vendas_carteira where cotacao = " & ID & " and Tipo = 'S' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    TotalServicos = IIf(IsNull(TBTotaisnota!TotalServicos), 0, TBTotaisnota!TotalServicos)
    Valor1 = IIf(IsNull(TBTotaisnota!Valor1), 0, TBTotaisnota!Valor1)
    If Valor1 > TotalServicos Then TotalServicos = Valor1
    Valor_PIS_Serv = IIf(IsNull(TBTotaisnota!Valor_PIS_Serv), 0, TBTotaisnota!Valor_PIS_Serv)
    Valor_Cofins_Serv = IIf(IsNull(TBTotaisnota!Valor_Cofins_Serv), 0, TBTotaisnota!Valor_Cofins_Serv)
    Valor_CSLL_Serv = IIf(IsNull(TBTotaisnota!Valor_CSLL_Serv), 0, TBTotaisnota!Valor_CSLL_Serv)
    TotalISS = IIf(IsNull(TBTotaisnota!TotalISS), 0, TBTotaisnota!TotalISS)
    Valor_INSS_Serv = IIf(IsNull(TBTotaisnota!Valor_INSS_Serv), 0, TBTotaisnota!Valor_INSS_Serv)
    Valor_IRPJ_Serv = IIf(IsNull(TBTotaisnota!Valor_IRPJ_Serv), 0, TBTotaisnota!Valor_IRPJ_Serv)
    Valor_IRRF_Serv = IIf(IsNull(TBTotaisnota!Valor_IRRF_Serv), 0, TBTotaisnota!Valor_IRRF_Serv)
    Valor_DAS = Valor_DAS + IIf(IsNull(TBTotaisnota!Valor_DAS), 0, TBTotaisnota!Valor_DAS)
End If
TBTotaisnota.Close

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from vendas_proposta where cotacao = " & ID, Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = True Then TBCarteira.AddNew
TBCarteira!dbl_Base_ICMS = Format(BASECALCULO, "###,##0.00")
TBCarteira!dbl_Valor_ICMS = Format(TotalICMS, "###,##0.00")
TBCarteira!dbl_Base_ICMS_Subst = Format(TotalBCICMSCST, "###,##0.00")
TBCarteira!dbl_Valor_ICMS_Subst = Format(TotalICMSCST, "###,##0.00")
TBCarteira!dbl_Valor_Total_Produtos = Format(TotalProduto, "###,##0.00")
TBCarteira!dbl_valor_total_servicos = Format(TotalServicos, "###,##0.00")
TBCarteira!TotalDesconto = Format(Valor_desconto, "###,##0.00") 'Format((TotalProduto - Valor_Desconto) + (TotalServicos - Valor1), "###,##0.00")
TBCarteira!dbl_Valor_Total_IPI = Format(SumIPI, "###,##0.00")

'Impostos produtos
TBCarteira!Total_PIS_prod = Format(Valor_PIS_Prod, "###,##0.00")
TBCarteira!Total_Cofins_prod = Format(Valor_Cofins_Prod, "###,##0.00")
TBCarteira!Total_CSLL_prod = Format(Valor_CSLL_Prod, "###,##0.00")
TBCarteira!Total_IRPJ_prod = Format(Valor_IRPJ_Prod, "###,##0.00")

'Impostos serviços
TBCarteira!Total_PIS_serv = Format(Valor_PIS_Serv, "###,##0.00")
TBCarteira!Total_Cofins_serv = Format(Valor_Cofins_Serv, "###,##0.00")
TBCarteira!Total_CSLL_serv = Format(Valor_CSLL_Serv, "###,##0.00")
TBCarteira!VlrTotaliss = Format(TotalISS)
TBCarteira!Total_INSS_serv = Format(Valor_INSS_Serv, "###,##0.00")
TBCarteira!Total_IRPJ_serv = Format(Valor_IRPJ_Serv, "###,##0.00")
TBCarteira!Total_IRRF_serv = Format(Valor_IRRF_Serv, "###,##0.00")

'===========================================================================
'Valor total do frete
TBCarteira!VTotalfrete = Format(Valor_total_Frete, "###,##0.00")
'===========================================================================

'SubTotal = Format(valor + Valor1 + Valor_total_Frete, "###,##0.00")
SubTotal = Format(valor + Valor1, "###,##0.00")
TBCarteira!SubTotal = Format(SubTotal, "###,##0.00")

'Impostos faturamento
TBCarteira!Total_DAS = Format(Valor_DAS, "###,##0.00")

'Retenção de PIS/Cofins
TBCarteira!Total_retencao_PIS = Format(Valor_Retencao_PIS, "###,##0.00")
TBCarteira!Total_retencao_Cofins = Format(Valor_Retencao_Cofins, "###,##0.00")

If Vendas_Proposta = True And TBCarteira!dbl_valor_total <> (SubTotal + SumIPI + TotalICMSCST) Then ProcSalvarPrevisaoPgto Format(SubTotal + SumIPI + TotalICMSCST, "###,##0.00")
'===============================================================================================
TBCarteira!dbl_valor_total = Format(SubTotal + SumIPI + TotalICMSCST + Valor_total_Frete, "###,##0.00")
'TBCarteira!dbl_valor_total = Format(SubTotal + SumIPI + TotalICMSCST, "###,##0.00")
'===============================================================================================
TBCarteira!Total_retorno = Format(VlrTotalRetorno, "###,##0.00")
TBCarteira.Update
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmitirPI()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtCotacao = "" Then
    USMsgBox ("Informe a proposta comercial antes de emitir o pedido interno."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente emitir pedido interno para esta proposta?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerifValidarAutomPropPI(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then
        If FunVerificaRegistroValidado("Vendas_proposta", "Cotacao = " & txtId, "mesma", "dessa proposta", "emitir PI", False, False) = False Then Exit Sub
    End If
    If txtStatus <> "ABERTA EM ANALISE" Then
        USMsgBox ("Só é permitido emitir pedido interno de proposta com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If txtIDcliente = "" Or txtIDcliente = "0" Then
        USMsgBox ("Só é permitido emitir pedido interno de proposta com o cliente cadastrado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from vendas_carteira where cotacao = " & txtId.Text & " and liberacao <> '" & "FATURADO" & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        Do While TBVendas.EOF = False
            'Verif. se todos os produtos/serviços estão cadastrados
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto where desenho = '" & TBVendas!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = True Then
                USMsgBox ("Não é permitido criar o pedido interno, pois os produtos/serviços precisam estar cadastrados."), vbExclamation, "CAPRIND v5.0"
                TBProduto.Close
                Exit Sub
            End If
            TBProduto.Close
            TBVendas.MoveNext
        Loop
        
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * FROM Clientes WHERE idcliente = " & txtIDcliente & " and Left(Tipo, 1) = 'J' and idTipoEmpresa = 1", Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = False Then
            If FunVerifRegimeTribCliForn(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
        End If
        TBClientes.Close
        
        PCCliente = InputBox("Favor informar o numero do pedido de compra do cliente.")
        If PCCliente = "" Then Exit Sub
                
        TextoFiltroUpdate = ""
        Permitido = False
        If FunVerifValidarAutomPropPI(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            Conexao.Execute "UPDATE vendas_proposta Set DtValidacao = '" & Now & "', RespValidacao = '" & pubUsuario & "', DtValidacaoPI = '" & Now & "', RespValidacaoPI = '" & pubUsuario & "' where Cotacao = " & txtId & " and DtValidacao IS NULL"
            Permitido = True
        End If
                
        TBVendas.MoveFirst
        Do While TBVendas.EOF = False
            TBVendas!Liberacao = "VENDIDA"
            TBVendas!Datavendas = Date
            If IsNull(TBVendas!prazofinaldias) = False And TBVendas!prazofinaldias <> "" Then TBVendas!PrazoFinal = FunDefinirPrazoPed(Date + TBVendas!prazofinaldias) Else TBVendas!PrazoFinal = Null
            TBVendas!Prazo_original = TBVendas!PrazoFinal
            TBVendas!PCCliente = Left(PCCliente, 15)
            txtDatavendas.Text = Format(TBVendas!Datavendas, "dd/mm/yyyy")
            TBVendas.Update
            
            If Permitido = True Then
                QuantSolicitado = TBVendas!Qtde_produzir
                ProcEmpenharProdEstoque Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, TBVendas!Desenho, True, False, TBVendas!Qtde_produzir
                If QuantSolicitado > 0 Then ProcEmpenharProdProduzindo Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBVendas!CODIGO, TBVendas!Desenho, TBVendas!PrazoFinal, True
            End If
            TBVendas.MoveNext
        Loop
    End If
    Conexao.Execute "UPDATE vendas_proposta Set Datavendas = '" & Date & "', Tipo = 'PRPE', Status = 'VENDIDA' where cotacao = " & txtId.Text
    USMsgBox ("Pedido interno gerado com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Emitir pedido interno"
    ID_documento = txtId
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    FunAtualizaStatusPropPI txtId
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista2 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista2)
        Lista.SetFocus
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_proposta where cotacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcPuxaDados
        ProcPuxaTotais
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifGravitem()
On Error GoTo tratar_erro

Permitido = False
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where desenho = '" & TBVendas!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    If USMsgBox("Deseja criar o produto com o código interno " & TBVendas!Desenho & " no cadastro de produtos/serviços?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        TBProduto.AddNew
        TBProduto!Classe = TBVendas!Familia
        TBProduto!CodManual = True
        TBProduto!Tipo = TBVendas!Tipo
        TBProduto!SubTipoItem = 1
        TBProduto!Producao = True
        TBProduto!Compras = False
        TBProduto!Vendas = True
        TBProduto!Qualidade = False
        TBProduto!Desenho = TBVendas!Desenho
        TBProduto!Data = Date
        TBProduto!Descricao = TBVendas!descricao_tecnica
        TBProduto!descricaotecnica = TBVendas!Descricao
        TBProduto!Observacoes = TBVendas!Observacoes
        TBProduto!Comprimento = TBVendas!Comprimento
        TBProduto!Largura = TBVendas!Largura
        TBProduto!Espessura = TBVendas!Espessura
        TBProduto!Dureza = TBVendas!Dureza
        TBProduto!Unidade = TBVendas!Unidade
        TBProduto!ID_CF = TBVendas!ID_CF
        TBProduto!RevDesenho = 0
        TBProduto!Responsavel = pubUsuario
        If txttipocliente <> "JR" And txttipocliente <> "FR" Then TBProduto!PConsumo = TBVendas!preco_unitario Else TBProduto!PRevenda = TBVendas!preco_unitario
        TBProduto.Update
    End If
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcpuxadadoslistaServicos()
On Error GoTo tratar_erro

txtcodservico = IIf(IsNull(TBProduto!Desenho), "", (TBProduto!Desenho))
NomeCampo = "Cidade onde foi executado o serviço"
If IsNull(TBProduto!Cidade) = False And TBProduto!Cidade <> "" Then Cmb_cidade_servico = TBProduto!Cidade
NomeCampo = "a unidade de estoque"
If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then txtunservico.Text = TBProduto!Unidade
NomeCampo = "a unidade comercial"
If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com_serv.Text = TBProduto!Unidade_com
NomeCampo = "a família"
If IsNull(TBProduto!Familia) = False And TBProduto!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbfamiliaservico.Text = TBProduto!Familia
    VerifDadosPadraoFamilia = True
End If
If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" Then
    NomeCampo = "o código de referência"
    cmbreferencia_serv.AddItem TBProduto!N_referencia
    cmbreferencia_serv = TBProduto!N_referencia
End If
1:
    txtid_servico.Text = TBProduto!CODIGO
    
    ProcCarregaDadosCFOPProdServ IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), False
    txtpcclienteserv = IIf(IsNull(TBProduto!PCCliente), "", TBProduto!PCCliente)
    Caminho_PC_serv_PI = IIf(IsNull(TBProduto!Caminho_PCCliente), "", TBProduto!Caminho_PCCliente)
    txtRev_serv = IIf(IsNull(TBProduto!Rev_codinterno), 0, TBProduto!Rev_codinterno)
    txtdescservico.Text = IIf(IsNull(TBProduto!descricao_tecnica), "", (TBProduto!descricao_tecnica))
    If Vendas_PI = True Then
        If IsNull(TBProduto!PrazoFinal) = False Then mskprazoservico = Format(TBProduto!PrazoFinal, "dd/mm/yyyy")
    Else
        txtPrazo_Servico = IIf(IsNull(TBProduto!prazofinaldias), "", TBProduto!prazofinaldias)
    End If
    txtdesccomservico.Text = IIf(IsNull(TBProduto!Descricao), "", (TBProduto!Descricao))
    txtqtservico.Text = IIf(IsNull(TBProduto!quantidade), "", (Format(TBProduto!quantidade, "###,##0.0000")))
    txtvlrunitservico.Text = IIf(IsNull(TBProduto!preco_unitario), "", (Format(TBProduto!preco_unitario, "###,##0.0000000000")))
    If TBProduto!Servico_cliente = True Then Chk_servico_executado_cliente.Value = 1 Else Chk_servico_executado_cliente.Value = 0
    txtiss.Text = IIf(IsNull(TBProduto!ISS), "", (TBProduto!ISS))
    txtvlrtotalservico.Text = IIf(IsNull(TBProduto!preco_lote), "", (Format(TBProduto!preco_lote, "###,##0.00")))
    
    If TBProduto!Desconto > 0 Then
        Chk_desc2.Value = 1
    Else
        Chk_desc2.Value = 0
        Chk_valor_desc2.Value = 0
    End If
    txtdesconto2.Text = IIf(IsNull(TBProduto!Desconto), "", TBProduto!Desconto)
    txtvalordesconto2.Text = IIf(IsNull(TBProduto!ValorDesconto), "", Format(TBProduto!ValorDesconto, "###,##0.0000000000"))
    txtvalorunitariodesc2 = IIf(IsNull(TBProduto!preco_unitario_desconto), "", Format(TBProduto!preco_unitario_desconto, "###,##0.0000000000"))
    
    txtObs_serv = IIf(IsNull(TBProduto!Observacoes), "", TBProduto!Observacoes)
    Txt_observacoes_fat_serv = IIf(IsNull(TBProduto!Obs_faturamento), "", TBProduto!Obs_faturamento)
    If TBProduto!Antecipacao_fat = True Then Chk_antecipacao_serv.Value = 1 Else Chk_antecipacao_serv.Value = 0
    If TBProduto!Faturamento_parcial = True Then Chk_faturamento_parcial_serv.Value = 1 Else Chk_faturamento_parcial_serv.Value = 0
    If TBProduto!Utiliza_mat_cons = True Then Chk_utiliza_mat_consignado_serv.Value = 1 Else Chk_utiliza_mat_consignado_serv.Value = 0
    IDAnalise_servico = IIf(IsNull(TBProduto!IDAnalise), 0, TBProduto!IDAnalise)
    txtComissaoServ.Text = IIf(IsNull(TBProduto!Comissao), "", TBProduto!Comissao)
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Nanalise from Vendas_analise where ID = " & IDAnalise_servico, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Txt_analise1 = IIf(IsNull(TBFIltro!Nanalise), "", TBFIltro!Nanalise)
    End If
    
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Valor_bloqueado from projproduto where Desenho = '" & TBProduto!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        If TBFIltro!Valor_bloqueado = True Then txtvlrunitservico.Locked = True Else txtvlrunitservico.Locked = False
    End If
    TBFIltro.Close
    'ProcBloqueiaLibera_Validacao
    
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste serviço."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosLista()
On Error GoTo tratar_erro

txtNomenclatura.Text = TBProduto!Desenho
NomeCampo = "a unidade de estoque"
If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun.Text = TBProduto!Unidade
NomeCampo = "a unidade comercial"
If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com.Text = TBProduto!Unidade_com
NomeCampo = "a família"
If IsNull(TBProduto!Familia) = False And TBProduto!Familia <> "" Then
    VerifDadosPadraoFamilia = False
    cmbfamilia.Text = TBProduto!Familia
    VerifDadosPadraoFamilia = True
End If
If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" Then
    NomeCampo = "o código de referência"
    cmbReferencia.AddItem TBProduto!N_referencia
    cmbReferencia = TBProduto!N_referencia
End If

ProcCarregaDadosCFOPProdServ IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), True

Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select Idclass, IDIntClasse from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBProduto!ID_CF), 0, TBProduto!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
If TBOSC.EOF = False Then
    Txt_ID_CF = TBOSC!Idclass
    Txt_CF = IIf(IsNull(TBOSC!IDIntClasse), "", TBOSC!IDIntClasse)
End If
TBOSC.Close

If IsNull(TBProduto!txt_CST) = False And TBProduto!txt_CST <> "" Then
    NomeCampo = "a CST"
CST:
    Cmb_CST_ICMS.Text = TBProduto!txt_CST
End If
    
1:
    txtid_produto.Text = TBProduto!CODIGO
    txtRev_cod = IIf(IsNull(TBProduto!Rev_codinterno), 0, TBProduto!Rev_codinterno)
    txtdesctecnica.Text = IIf(IsNull(TBProduto!descricao_tecnica), "", (TBProduto!descricao_tecnica))
    txtEspecificacoes.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
'=============================================================================================
    txtvFrete = IIf(IsNull(TBProduto!vFrete), "0", Format(TBProduto!vFrete, "###,##0.00"))
'=============================================================================================
    Txt_observacoes_prod = IIf(IsNull(TBProduto!Observacoes), "", TBProduto!Observacoes)
    Txt_observacoes_fat_prod = IIf(IsNull(TBProduto!Obs_faturamento), "", TBProduto!Obs_faturamento)
    txtinspecao = IIf(IsNull(TBProduto!Inspecao), "", TBProduto!Inspecao)
    txtembalagem = IIf(IsNull(TBProduto!Embalagem), "", TBProduto!Embalagem)
    txtGravacao = IIf(IsNull(TBProduto!Gravacao), "", TBProduto!Gravacao)
    If IsNull(TBProduto!Prioridade) = False And TBProduto!Prioridade <> "" Then Cmb_prioridade = TBProduto!Prioridade
    If TBProduto!Antecipacao_fat = True Then Chk_antecipacao.Value = 1 Else Chk_antecipacao.Value = 0
    If TBProduto!Faturamento_parcial = True Then Chk_faturamento_parcial.Value = 1 Else Chk_faturamento_parcial.Value = 0
    If TBProduto!Novo_projeto = True Then chkNovo_projeto.Value = 1 Else chkNovo_projeto.Value = 0
    If TBProduto!Utiliza_mat_cons = True Then Chk_utiliza_mat_consignado.Value = 1 Else Chk_utiliza_mat_consignado.Value = 0
    txtespessura = IIf(IsNull(TBProduto!Espessura), "", TBProduto!Espessura)
    txtLargura = IIf(IsNull(TBProduto!Largura), "", TBProduto!Largura)
    txtComprimento = IIf(IsNull(TBProduto!Comprimento), "", TBProduto!Comprimento)
    txtDureza = IIf(IsNull(TBProduto!Dureza), "", TBProduto!Dureza)
    txtpccliente = IIf(IsNull(TBProduto!PCCliente), "", TBProduto!PCCliente)
    Caminho_PC_prod_PI = IIf(IsNull(TBProduto!Caminho_PCCliente), "", TBProduto!Caminho_PCCliente)
    Txt_n_serie = IIf(IsNull(TBProduto!N_Serie), "", TBProduto!N_Serie)
    Txt_n_item = IIf(IsNull(TBProduto!N_item), "", TBProduto!N_item)
    If Vendas_PI = True Then
        If IsNull(TBProduto!PrazoFinal) = False Then mskprazo.Text = Format(TBProduto!PrazoFinal, "dd/mm/yyyy")
    Else
        txtPrazo_Produto.Text = IIf(IsNull(TBProduto!prazofinaldias), "", TBProduto!prazofinaldias)
    End If
    txtQuantidade.Text = IIf(IsNull(TBProduto!quantidade), "", Format(TBProduto!quantidade, "###,##0.0000"))
    txtvalorunitario.Text = Format(TBProduto!preco_unitario, "###,##0.0000000000")
    txtInt_ipi = IIf(IsNull(TBProduto!int_IPI), "", (TBProduto!int_IPI))
    txtdbl_valoripi.Text = IIf(IsNull(TBProduto!dbl_valoripi), "", Format(TBProduto!dbl_valoripi, "###,##0.00"))
    txtint_icms = IIf(IsNull(TBProduto!IntICMS), "", (TBProduto!IntICMS))
    txtvalor_icms = IIf(IsNull(TBProduto!dbl_Valor_ICMS), "", Format(TBProduto!dbl_Valor_ICMS, "###,##0.00"))
    txtvalor_total.Text = IIf(IsNull(TBProduto!preco_lote), "", (Format(TBProduto!preco_lote, "###,##0.00")))
    If TBProduto!retorno = True Then chkRetorno.Value = 1 Else chkRetorno.Value = 0
    Txt_data_retorno = IIf(IsNull(TBProduto!Data_retorno), "__/__/____", TBProduto!Data_retorno)
    
    If TBProduto!Desconto > 0 Then
        Chk_desc.Value = 1
    Else
        Chk_desc.Value = 0
        Chk_valor_desc.Value = 0
    End If
    txtDesconto.Text = IIf(IsNull(TBProduto!Desconto), "", TBProduto!Desconto)
    txtvalordesconto.Text = IIf(IsNull(TBProduto!ValorDesconto), "", Format(TBProduto!ValorDesconto, "###,##0.0000000000"))
    txtvalorunitariodesc = IIf(IsNull(TBProduto!preco_unitario_desconto), "", Format(TBProduto!preco_unitario_desconto, "###,##0.0000000000"))
    txtvFrete = IIf(IsNull(TBProduto!vFrete), "", Format(TBProduto!vFrete, "###,##0.00"))
   
    txtComissao.Text = IIf(IsNull(TBProduto!Comissao), "", TBProduto!Comissao)
    IDAnalise = IIf(IsNull(TBProduto!IDAnalise), 0, TBProduto!IDAnalise)
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Nanalise from Vendas_analise where ID = " & IDAnalise, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Txt_analise = IIf(IsNull(TBFIltro!Nanalise), "", TBFIltro!Nanalise)
    End If
    
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select Valor_bloqueado from projproduto where Desenho = '" & TBProduto!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        If TBFIltro!Valor_bloqueado = True Then txtvalorunitario.Locked = True Else txtvalorunitario.Locked = False
    End If
    TBFIltro.Close
    'ProcBloqueiaLibera_Validacao
    
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        If NomeCampo = "a CST" Then
            Cmb_CST_ICMS.AddItem TBProduto!txt_CST
            GoTo CST
        Else
            USMsgBox ("Não foi encontrado " & NomeCampo & " deste produto."), vbExclamation, "CAPRIND v5.0"
            GoTo 1
        End If
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosCFOPProdServ(ID_CFOP As Long, Prod As Boolean)
On Error GoTo tratar_erro

Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select IDCountCfop, ID_CFOP, Txt_descricao from tbl_NaturezaOperacao where IDCountCfop = " & ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBOSC.EOF = False Then
    If Prod = True Then
        Txt_ID_CFOP_prod = TBOSC!IDCountCfop
        Txt_CFOP_prod = IIf(IsNull(TBOSC!ID_CFOP), "", TBOSC!ID_CFOP)
        Txt_natureza_operacao_prod = IIf(IsNull(TBOSC!Txt_descricao), "", TBOSC!Txt_descricao)
    Else
        Txt_ID_CFOP_serv = TBOSC!IDCountCfop
        Txt_CFOP_serv = IIf(IsNull(TBOSC!ID_CFOP), "", TBOSC!ID_CFOP)
        Txt_natureza_operacao_serv = IIf(IsNull(TBOSC!Txt_descricao), "", TBOSC!Txt_descricao)
    End If
End If
TBOSC.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalistaProdutos(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros1.Caption = "Nº de registros: 0"
lblPaginas1.Caption = "Página: 0 de: 0"
Listprod.ListItems.Clear

If Vendas_Proposta = True Then
TextoFiltro = "Select * from vendas_carteira where cotacao = " & txtId & " and TIPO = 'P' order by Codigo"
Else
TextoFiltro = "Select VC.* from vendas_carteira VC INNER JOIN vendas_proposta VP on VP.Cotacao = VC.Cotacao where VP.cotacao = " & txtId & " and VC.Tipo = 'P' and (VP.Tipo = 'PE' or VP.Tipo = 'PRPE') and (VC.liberacao = 'VENDIDA' or VC.liberacao = 'REVISADA' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL' or VC.liberacao = 'CANCELADO') order by VC.Codigo"
End If

Set TBLISTA_Vendas_PI1 = CreateObject("adodb.recordset")
TBLISTA_Vendas_PI1.Open TextoFiltro, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_PI1.EOF = False Then ProcExibePagina1 (Pagina)
ProcGravarTotais txtId
ProcPuxaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina1(Pagina)
On Error GoTo tratar_erro

Listprod.ListItems.Clear
TBLISTA_Vendas_PI1.PageSize = IIf(txtNreg1 = "", 30, txtNreg1)
TBLISTA_Vendas_PI1.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_PI1.PageSize
ContadorReg = 1
'PBLista.Min = 0
'PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_PI1.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_PI1.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_PI1.PageSize)
'PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendas_PI1.EOF = False And (ContadorReg <= TamanhoPagina)
    With Listprod.ListItems
        .Add , , IIf(IsNull(TBLISTA_Vendas_PI1!CODIGO), "", TBLISTA_Vendas_PI1!CODIGO)
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_PI1!Desenho), "", TBLISTA_Vendas_PI1!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Vendas_PI1!descricao_tecnica), "", TBLISTA_Vendas_PI1!descricao_tecnica)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendas_PI1!quantidade), "", Format(TBLISTA_Vendas_PI1!quantidade, "###,##0.0000"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_PI1!preco_unitario), "", Format(TBLISTA_Vendas_PI1!preco_unitario, "###,##0.0000000000"))
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Vendas_PI1!Desconto), "", TBLISTA_Vendas_PI1!Desconto)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Vendas_PI1!ValorDesconto), "", Format(TBLISTA_Vendas_PI1!ValorDesconto, "###,##0.0000000000"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Vendas_PI1!preco_unitario_desconto), "", Format(TBLISTA_Vendas_PI1!preco_unitario_desconto, "###,##0.0000000000"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Vendas_PI1!preco_lote), "", Format(TBLISTA_Vendas_PI1!preco_lote, "###,##0.00"))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Vendas_PI1!Liberacao), "", TBLISTA_Vendas_PI1!Liberacao)
        If Vendas_PI = True Then PrazoFinalTexto = IIf(IsNull(TBLISTA_Vendas_PI1!PrazoFinal), "", Format(TBLISTA_Vendas_PI1!PrazoFinal, "dd/mm/yy")) Else PrazoFinalTexto = IIf(IsNull(TBLISTA_Vendas_PI1!prazofinaldias), "", TBLISTA_Vendas_PI1!prazofinaldias & " dias")
        .Item(.Count).SubItems(10) = PrazoFinalTexto
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Vendas_PI1!PCCliente), "", TBLISTA_Vendas_PI1!PCCliente)
    End With
    TBLISTA_Vendas_PI1.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    'PBLista.Value = Contador
Loop
lblRegistros1.Caption = "Nº de registros: " & TBLISTA_Vendas_PI1.RecordCount
If TBLISTA_Vendas_PI1.AbsolutePage = adPosBOF Then
   lblPaginas1.Caption = "Página: 1 de: " & TBLISTA_Vendas_PI1.PageCount
ElseIf TBLISTA_Vendas_PI1.AbsolutePage = adPosEOF Then
        lblPaginas1.Caption = "Página: " & TBLISTA_Vendas_PI1.PageCount & " de: " & TBLISTA_Vendas_PI1.PageCount
    Else
        lblPaginas1.Caption = "Página: " & TBLISTA_Vendas_PI1.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_PI1.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalistaServicos(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros2.Caption = "Nº de registros: 0"
lblPaginas2.Caption = "Página: 0 de: 0"
ListaServicos.ListItems.Clear
If Vendas_Proposta = True Then TextoFiltro = "Select * from vendas_carteira where cotacao = " & txtId & " and TIPO = 'S' order by Codigo" Else TextoFiltro = "Select VC.* from vendas_carteira VC INNER JOIN vendas_proposta VP on VP.Cotacao = VC.Cotacao where VP.cotacao = " & txtId & " and VC.Tipo = 'S' and (VP.Tipo = 'PE' or VP.Tipo = 'PRPE') and (VC.liberacao = 'VENDIDA' or VC.liberacao = 'REVISADA' or VC.liberacao = 'FATURAR' or VC.liberacao = 'FATURAR PARCIAL' or VC.liberacao = 'FATURADO' or VC.liberacao = 'FATURADO PARCIAL' or VC.liberacao = 'CANCELADO') order by VC.Codigo"
Set TBLISTA_Vendas_PI2 = CreateObject("adodb.recordset")
TBLISTA_Vendas_PI2.Open TextoFiltro, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_PI2.EOF = False Then ProcExibePagina2 (Pagina)
ProcGravarTotais txtId
ProcPuxaTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina2(Pagina)
On Error GoTo tratar_erro

ListaServicos.ListItems.Clear
TBLISTA_Vendas_PI2.PageSize = IIf(txtNreg2 = "", 30, txtNreg2)
TBLISTA_Vendas_PI2.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_PI2.PageSize
ContadorReg = 1
'PBLista.Min = 0
'PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_PI2.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_PI2.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_PI2.PageSize)
'PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendas_PI2.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaServicos.ListItems
        .Add , , IIf(IsNull(TBLISTA_Vendas_PI2!CODIGO), "", TBLISTA_Vendas_PI2!CODIGO)
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_PI2!Desenho), "", TBLISTA_Vendas_PI2!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Vendas_PI2!descricao_tecnica), "", TBLISTA_Vendas_PI2!descricao_tecnica)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendas_PI2!quantidade), "", Format(TBLISTA_Vendas_PI2!quantidade, "###,##0.0000"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_PI2!preco_unitario), "", Format(TBLISTA_Vendas_PI2!preco_unitario, "###,##0.0000000000"))
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Vendas_PI2!Desconto), "", TBLISTA_Vendas_PI2!Desconto)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Vendas_PI2!ValorDesconto), "", Format(TBLISTA_Vendas_PI2!ValorDesconto, "###,##0.0000000000"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Vendas_PI2!preco_unitario_desconto), "", Format(TBLISTA_Vendas_PI2!preco_unitario_desconto, "###,##0.0000000000"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Vendas_PI2!preco_lote), "", Format(TBLISTA_Vendas_PI2!preco_lote, "###,##0.00"))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Vendas_PI2!Liberacao), "", TBLISTA_Vendas_PI2!Liberacao)
        If Vendas_PI = True Then PrazoFinalTexto = IIf(IsNull(TBLISTA_Vendas_PI2!PrazoFinal), "", Format(TBLISTA_Vendas_PI2!PrazoFinal, "dd/mm/yy")) Else PrazoFinalTexto = IIf(IsNull(TBLISTA_Vendas_PI2!prazofinaldias), "", TBLISTA_Vendas_PI2!prazofinaldias & " dias")
        .Item(.Count).SubItems(10) = PrazoFinalTexto
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Vendas_PI2!PCCliente), "", TBLISTA_Vendas_PI2!PCCliente)
    End With
    TBLISTA_Vendas_PI2.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    'PBLista.Value = Contador
Loop
lblRegistros2.Caption = "Nº de registros: " & TBLISTA_Vendas_PI2.RecordCount
If TBLISTA_Vendas_PI2.AbsolutePage = adPosBOF Then
   lblPaginas2.Caption = "Página: 1 de: " & TBLISTA_Vendas_PI2.PageCount
ElseIf TBLISTA_Vendas_PI2.AbsolutePage = adPosEOF Then
        lblPaginas2.Caption = "Página: " & TBLISTA_Vendas_PI2.PageCount & " de: " & TBLISTA_Vendas_PI2.PageCount
    Else
        lblPaginas2.Caption = "Página: " & TBLISTA_Vendas_PI2.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_PI2.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparComercial()
On Error GoTo tratar_erro

txtcalculos.Text = "N/A"
txtimpostos.Text = "N/A"
txtCondicoes.Text = "N/A"
txtgarantia.Text = "N/A"
txtObservacoes.Text = ""
txtReajuste.Text = "N/A"
txttransporte.Text = "N/A"
txtValidade.Text = "N/A"
Txt_ID_entrega = 0
txtlocal_entrega.Clear
Txt_ID_cobranca = 0
txtlocal_cobranca.Clear
'Cmb_tipo_transp.ListIndex = -1
'cmbtransportadora.ListIndex = -1
'Cmb_tipo_transp2.ListIndex = -1
'cmbtransportadora2.ListIndex = -1
cmbMoeda.ListIndex = -1
Txt_valor_moeda = ""
'txt_ValorFrete = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosComercial()
On Error GoTo tratar_erro

ProcLimparComercial
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM vendas_comercial WHERE cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    If TBCotacao!analize = "Sim" Or TBCotacao!analize = "Não" Then txtAnalize.Text = TBCotacao!analize
    txtcalculos = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
    txtimpostos = IIf(IsNull(TBCotacao!impostos), "", TBCotacao!impostos)
    txtCondicoes = IIf(IsNull(TBCotacao!condicoes), "", TBCotacao!condicoes)
    txtgarantia = IIf(IsNull(TBCotacao!garantia), "", TBCotacao!garantia)
    txtObservacoes = IIf(IsNull(TBCotacao!Observacoes), "", TBCotacao!Observacoes)
    txtReajuste = IIf(IsNull(TBCotacao!reajuste), "", TBCotacao!reajuste)
    txttransporte = IIf(IsNull(TBCotacao!transporte), "", TBCotacao!transporte)
    txtValidade = IIf(IsNull(TBCotacao!validade), "", TBCotacao!validade)
    txtcalculos = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
    
'===========================================================================================
        If IsNull(TBCotacao!Tipo_transp2) = False And TBCotacao!Tipo_transp2 <> "" Then
        Select Case TBCotacao!Tipo_transp2
            Case "C": txtTipoTransp(1).Text = "Cliente"
            Case "F": txtTipoTransp(1).Text = "Fornecedor"
            Case "E": txtTipoTransp(1).Text = "Empresa"
        End Select
    End If
    NomeCampo = "a transportadora"
    
    If IsNull(TBCotacao!Redespacho) = False And TBCotacao!Redespacho <> "" Then txtRedespacho = TBCotacao!Redespacho
    If IsNull(TBCotacao!Tipo_Frete) = False And TBCotacao!Tipo_Frete <> "" Then cmb_Tipo_Frete.Text = TBCotacao!Tipo_Frete
    
        If IsNull(TBCotacao!Tipo_transp) = False And TBCotacao!Tipo_transp <> "" Then
        Select Case TBCotacao!Tipo_transp
            Case "C": txtTipoTransp(0).Text = "Cliente"
            Case "F": txtTipoTransp(0).Text = "Fornecedor"
            Case "E": txtTipoTransp(0).Text = "Empresa"
        End Select
    End If
    
'===========================================================================================
    
    With txtlocal_entrega
        .AddItem ""
        If IsNull(TBCotacao!Local_entrega) = False And TBCotacao!Local_entrega <> "" Then
            .AddItem TBCotacao!Local_entrega
            .Text = TBCotacao!Local_entrega
            Txt_ID_entrega = IIf(IsNull(TBCotacao!ID_entrega), 0, TBCotacao!ID_entrega)
        End If
    End With
    With txtlocal_cobranca
        .AddItem ""
        If IsNull(TBCotacao!Local_Cobranca) = False And TBCotacao!Local_Cobranca <> "" Then
            .AddItem TBCotacao!Local_Cobranca
            .Text = TBCotacao!Local_Cobranca
            Txt_ID_cobranca = IIf(IsNull(TBCotacao!ID_Cobranca), 0, TBCotacao!ID_Cobranca)
        End If
    End With

    If IsNull(TBCotacao!Tipo_transp) = False And TBCotacao!Tipo_transp <> "" Then
        Select Case TBCotacao!Tipo_transp
            Case "C": txtTipoTransp(0).Text = "Cliente"
            Case "F": txtTipoTransp(0).Text = "Fornecedor"
            Case "E": txtTipoTransp(0).Text = "Empresa"
        End Select
    End If
    NomeCampo = "a transportadora"
    If IsNull(TBCotacao!Transportadora) = False And TBCotacao!Transportadora <> "" Then
    txtTransportadora = TBCotacao!Transportadora
    txtidTransportadora.Text = TBCotacao!IdIntTransp
    End If
    
    If IsNull(TBCotacao!Moeda) = False And TBCotacao!Moeda <> "" Then cmbMoeda = TBCotacao!Moeda
    Txt_valor_moeda = IIf(IsNull(TBCotacao!Valor_moeda), "", Format(TBCotacao!Valor_moeda, "###,##0.0000"))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAdicionarCliente_Click()
On Error GoTo tratar_erro
    
ProcLocalizarCliente

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
If txtCotacao = "" Then
    USMsgBox ("Informe " & IIf(Vendas_Proposta = True, "a proposta comercial", "o pedido interno") & " antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar " & IIf(Vendas_Proposta = True, "a proposta comercial ", "o pedido interno ") & txtCotacao.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Revisar = False
    ProcCopiarPI_Proposta
    '==================================
    Modulo = Formulario
    Evento = "Novo"
    ID_documento = txtId
    Documento = IIf(Vendas_Proposta = True, "Nº proposta: ", "Nº pedido: ") & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Frame1(1).Enabled = True
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiarPI_Proposta()
On Error GoTo tratar_erro

IDAntigo = txtId.Text
Set TBProposta = CreateObject("adodb.recordset")
TBProposta.Open "Select * from vendas_proposta where cotacao = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBProposta.EOF = False Then
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from vendas_proposta where Year(Data) = '" & Year(Date) & "' order by Ordenarproposta", Conexao, adOpenKeyset, adLockOptimistic
    If TBCotacao.EOF = False Then
        TBCotacao.MoveLast
        If Len(TBCotacao!Ncotacao) = 7 Then
            Cotacao = Left(TBCotacao("ncotacao"), Len(TBCotacao!Ncotacao) - 3) + 1
        End If
        If Len(TBCotacao!Ncotacao) = 8 Then
            Cotacao = Left(TBCotacao("ncotacao"), Len(TBCotacao!Ncotacao) - 3) + 1
        End If
    Else
        Cotacao = 1
    End If
    txtrevisao.Text = 0
    If Len(Cotacao) = 5 Then txtCotacao.Text = Cotacao & "/" & Right(Year(Date), 2)
    If Len(Cotacao) = 4 Then txtCotacao.Text = Cotacao & "/" & Right(Year(Date), 2)
    If Len(Cotacao) = 3 Then txtCotacao.Text = "0" & Cotacao & "/" & Right(Year(Date), 2)
    If Len(Cotacao) = 2 Then txtCotacao.Text = "00" & Cotacao & "/" & Right(Year(Date), 2)
    If Len(Cotacao) = 1 Then txtCotacao.Text = "000" & Cotacao & "/" & Right(Year(Date), 2)
    txt_dataelaborado.Value = Date
    TBCotacao.AddNew
    TBCotacao!Ncotacao = txtCotacao.Text
    TBCotacao!Data = Date
    If TBProposta!Revisao <> "" Then TBCotacao!Revisao = 0
    TBCotacao!Responsavel = pubUsuario
    ProcCopiarInfGerais
    ProcCopiarTotais
    TBCotacao.Update
    txtId.Text = TBCotacao!Cotacao
    Conexao.Execute "Update vendas_proposta Set ordenarproposta = " & txtId & " where cotacao = " & txtId

    TBCotacao.Close
    
    ProcCopiarDetalhes
    ProcCopiarItem
    ProcCopiarServico
End If
TBProposta.Close

USMsgBox (IIf(Vendas_Proposta = True, "Proposta copiada", "Pedido interno copiado") & " com sucesso."), vbInformation, "CAPRIND v5.0"
Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select VP.*, CL.CPF_CNPJ as CNPJ_CPF, CL.CEP as CEP, CL.RG_IE from vendas_proposta VP inner join Clientes CL on VP.IDcliente = CL.IDCliente where cotacao ="
TBAbrir.Open StrSql & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then ProcPuxaDados
TBAbrir.Close
ProcAtualizalistaProdutos (1)
ProcAtualizalistaServicos (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarInfGerais()
On Error GoTo tratar_erro

TBCotacao!ID_empresa = TBProposta!ID_empresa
If Vendas_Proposta = True Then
    TBCotacao!status = "ABERTA EM ANALISE"
    TBCotacao!Tipo = "PR"
Else
    TBCotacao!status = "VENDIDA"
    TBCotacao!Datavendas = Date
    TBCotacao!Tipo = "PE"
End If
TBCotacao!VI = TBProposta!VI
TBCotacao!VE = TBProposta!VE
TBCotacao!regiao = TBProposta!regiao
TBCotacao!IDCliente = TBProposta!IDCliente
TBCotacao!Cliente = TBProposta!Cliente
TBCotacao!Remetente = TBProposta!Remetente
TBCotacao!Referente = TBProposta!Referente
TBCotacao!Fax = TBProposta!Fax
TBCotacao!Email = TBProposta!Email

TBCotacao!Tipo_endereco = TBProposta!Tipo_endereco
TBCotacao!Endereco = TBProposta!Endereco
TBCotacao!Numero = TBProposta!Numero
TBCotacao!complemento = TBProposta!complemento
TBCotacao!Tipo_bairro = TBProposta!Tipo_bairro
TBCotacao!Bairro = TBProposta!Bairro
TBCotacao!Cidade = TBProposta!Cidade

TBCotacao!telefone = TBProposta!telefone
TBCotacao!Departamento = TBProposta!Departamento
TBCotacao!UF = TBProposta!UF
TBCotacao!Tipo_cliente = TBProposta!Tipo_cliente
TBCotacao!Obs = TBProposta!Obs
TBCotacao!Ref = TBProposta!Ref
TBCotacao!Vend_ext = TBProposta!Vend_ext
TBCotacao!vend_int = TBProposta!vend_int
TBCotacao!Regime = FunVerifRegimeEmpresa(TBProposta!ID_empresa)
If TBCotacao!Regime = 1 Then TBCotacao!TabelaSN = TBProposta!TabelaSN

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarDetalhes()
On Error GoTo tratar_erro

Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_comercial where cotacao = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_comercial", Conexao, adOpenKeyset, adLockOptimistic
    TBAbrir.AddNew
    TBAbrir!analize = TBVendas!analize
    TBAbrir!calculos = TBVendas!calculos
    TBAbrir!impostos = TBVendas!impostos
    TBAbrir!condicoes = TBVendas!condicoes
    TBAbrir!Cotacao = txtId.Text
    TBAbrir!garantia = TBVendas!garantia
    TBAbrir!Observacoes = TBVendas!Observacoes
    TBAbrir!reajuste = TBVendas!reajuste
    TBAbrir!transporte = TBVendas!transporte
    TBAbrir!validade = TBVendas!validade
    TBAbrir!ID_entrega = TBVendas!ID_entrega
    TBAbrir!Local_entrega = TBVendas!Local_entrega
    TBAbrir!ID_Cobranca = TBVendas!ID_Cobranca
    TBAbrir!Local_Cobranca = TBVendas!Local_Cobranca
    TBAbrir!Tipo_transp = TBVendas!Tipo_transp
    TBAbrir!Transportadora = TBVendas!Transportadora
    TBAbrir!IdIntTransp = TBVendas!IdIntTransp
    TBAbrir!Escopo_fornecimento = TBVendas!Escopo_fornecimento
    TBAbrir!Moeda = TBVendas!Moeda
    TBAbrir!Valor_moeda = TBVendas!Valor_moeda
    
    '======================================================================================
    TBAbrir!Tipo_Frete = TBVendas!Tipo_Frete
    TBAbrir!Tipo_transp2 = TBVendas!Tipo_transp2
    TBAbrir!Redespacho = TBVendas!Redespacho
    '======================================================================================
    
    TBAbrir.Update
    TBAbrir.Close
End If
TBVendas.Close

Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_proposta_previsaopgto where cotacao = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    Do While TBVendas.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from vendas_proposta_previsaopgto", Conexao, adOpenKeyset, adLockOptimistic
        TBAbrir.AddNew
        TBAbrir!Cotacao = txtId.Text
        TBAbrir!Data = TBVendas!Data
        TBAbrir!valor = TBVendas!valor
        TBAbrir!Parcela = TBVendas!Parcela
        TBAbrir.Update
        TBAbrir.Close
        TBVendas.MoveNext
    Loop
End If
TBVendas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarItem()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_carteira where cotacao = " & IDAntigo & " and tipo = 'P' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from vendas_carteira", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBLISTA.EOF = False
        TBProduto.AddNew
        TBProduto!Tem_ordem = False
        TBProduto!Cotacao = txtId.Text
        If IsNull(TBLISTA!Desenho) = False Then TBProduto!Desenho = TBLISTA!Desenho
        If IsNull(TBLISTA!N_referencia) = False Then TBProduto!N_referencia = TBLISTA!N_referencia
        If IsNull(TBLISTA!ID_CFOP) = False Then TBProduto!ID_CFOP = TBLISTA!ID_CFOP
        If IsNull(TBLISTA!Rev_codinterno) = False Then TBProduto!Rev_codinterno = TBLISTA!Rev_codinterno
        If IsNull(TBLISTA!descricao_tecnica) = False Then TBProduto!descricao_tecnica = TBLISTA!descricao_tecnica
        If IsNull(TBLISTA!Descricao) = False Then TBProduto!Descricao = TBLISTA!Descricao
        If IsNull(TBLISTA!quantidade) = False Then TBProduto!quantidade = TBLISTA!quantidade
        If IsNull(TBLISTA!Desconto) = False Then TBProduto!Desconto = TBLISTA!Desconto
        If IsNull(TBLISTA!ValorDesconto) = False Then TBProduto!ValorDesconto = TBLISTA!ValorDesconto
        If IsNull(TBLISTA!preco_unitario_desconto) = False Then TBProduto!preco_unitario_desconto = TBLISTA!preco_unitario_desconto
        If IsNull(TBLISTA!preco_unitario) = False Then TBProduto!preco_unitario = TBLISTA!preco_unitario
        If IsNull(TBLISTA!preco_lote) = False Then TBProduto!preco_lote = TBLISTA!preco_lote
        If IsNull(TBLISTA!Comprimento) = False Then TBProduto!Comprimento = TBLISTA!Comprimento
        If IsNull(TBLISTA!Largura) = False Then TBProduto!Largura = TBLISTA!Largura
        If IsNull(TBLISTA!Espessura) = False Then TBProduto!Espessura = TBLISTA!Espessura
        If IsNull(TBLISTA!Dureza) = False Then TBProduto!Dureza = TBLISTA!Dureza
        If IsNull(TBLISTA!ID_CF) = False Then TBProduto!ID_CF = TBLISTA!ID_CF
        If IsNull(TBLISTA!Unidade) = False Then TBProduto!Unidade = TBLISTA!Unidade
        If IsNull(TBLISTA!Unidade_com) = False Then TBProduto!Unidade_com = TBLISTA!Unidade_com
        If IsNull(TBLISTA!Familia) = False Then TBProduto!Familia = TBLISTA!Familia
        If IsNull(TBLISTA!N_Serie) = False Then TBProduto!N_Serie = TBLISTA!N_Serie
        If IsNull(TBLISTA!PrazoFinal) = False Then TBProduto!PrazoFinal = TBLISTA!PrazoFinal
        If IsNull(TBLISTA!Prazo_original) = False Then TBProduto!Prazo_original = TBLISTA!Prazo_original
        If IsNull(TBLISTA!prazofinaldias) = False Then TBProduto!prazofinaldias = TBLISTA!prazofinaldias
        If IsNull(TBLISTA!Comissao) = False Then TBProduto!Comissao = TBLISTA!Comissao
        If IsNull(TBLISTA!ValorComissao) = False Then TBProduto!ValorComissao = TBLISTA!ValorComissao
        If IsNull(TBLISTA!Qtde_produzir) = False Then TBProduto!Qtde_produzir = TBLISTA!Qtde_produzir
                
        'Impostos
        If IsNull(TBLISTA!IntICMS) = False Then TBProduto!IntICMS = TBLISTA!IntICMS
        If IsNull(TBLISTA!dbl_Valor_ICMS) = False Then TBProduto!dbl_Valor_ICMS = TBLISTA!dbl_Valor_ICMS
        If IsNull(TBLISTA!int_IPI) = False Then TBProduto!int_IPI = TBLISTA!int_IPI
        If IsNull(TBLISTA!dbl_valoripi) = False Then TBProduto!dbl_valoripi = TBLISTA!dbl_valoripi
        If IsNull(TBLISTA!PIS_Prod) = False Then TBProduto!PIS_Prod = TBLISTA!PIS_Prod
        If IsNull(TBLISTA!Total_PIS_prod) = False Then TBProduto!Total_PIS_prod = TBLISTA!Total_PIS_prod
        If IsNull(TBLISTA!Cofins_Prod) = False Then TBProduto!Cofins_Prod = TBLISTA!Cofins_Prod
        If IsNull(TBLISTA!Total_Cofins_prod) = False Then TBProduto!Total_Cofins_prod = TBLISTA!Total_Cofins_prod
        If IsNull(TBLISTA!CSLL_Prod) = False Then TBProduto!CSLL_Prod = TBLISTA!CSLL_Prod
        If IsNull(TBLISTA!Total_CSLL_prod) = False Then TBProduto!Total_CSLL_prod = TBLISTA!Total_CSLL_prod
        If IsNull(TBLISTA!IRPJ_Prod) = False Then TBProduto!IRPJ_Prod = TBLISTA!IRPJ_Prod
        If IsNull(TBLISTA!Total_IRPJ_prod) = False Then TBProduto!Total_IRPJ_prod = TBLISTA!Total_IRPJ_prod
        If IsNull(TBLISTA!cpp) = False Then TBProduto!cpp = TBLISTA!cpp
        If IsNull(TBLISTA!Total_CPP) = False Then TBProduto!Total_CPP = TBLISTA!Total_CPP
        If IsNull(TBLISTA!DAS) = False Then TBProduto!DAS = TBLISTA!DAS
        If IsNull(TBLISTA!Total_DAS) = False Then TBProduto!Total_DAS = TBLISTA!Total_DAS
        If IsNull(TBLISTA!txt_CST) = False Then TBProduto!txt_CST = TBLISTA!txt_CST
        If IsNull(TBLISTA!BC_ICMS) = False Then TBProduto!BC_ICMS = TBLISTA!BC_ICMS
        If IsNull(TBLISTA!BC_ICMS_ST) = False Then TBProduto!BC_ICMS_ST = TBLISTA!BC_ICMS_ST
        If IsNull(TBLISTA!Valor_ICMS_ST) = False Then TBProduto!Valor_ICMS_ST = TBLISTA!Valor_ICMS_ST
        TBProduto!Valor_Retencao_PIS = IIf(IsNull(TBLISTA!Valor_Retencao_PIS), 0, TBLISTA!Valor_Retencao_PIS)
        TBProduto!Valor_Retencao_Cofins = IIf(IsNull(TBLISTA!Valor_Retencao_Cofins), 0, TBLISTA!Valor_Retencao_Cofins)
                
        If Vendas_Proposta = True Then
            TBProduto!Liberacao = "ABERTA EM ANALISE"
        Else
            TBProduto!Liberacao = "VENDIDA"
            TBProduto!Datavendas = Date
        End If
        TBProduto!Tipo = "P"
        TBProduto!retorno = TBLISTA!retorno
        TBProduto!Observacoes = TBLISTA!Observacoes
        TBProduto!Obs_faturamento = TBLISTA!Obs_faturamento
        TBProduto!Antecipacao_fat = TBLISTA!Antecipacao_fat
        TBProduto!Faturamento_parcial = TBLISTA!Faturamento_parcial
        TBProduto!Inspecao = TBLISTA!Inspecao
        TBProduto!Embalagem = TBLISTA!Embalagem
        TBProduto!Gravacao = TBLISTA!Gravacao
        TBProduto!Novo_projeto = TBLISTA!Novo_projeto
        TBProduto!Prioridade = TBLISTA!Prioridade
        
        If Revisar = True Then
            TBLISTA!Liberacao = "REVISADA"
            TBLISTA.Update
        End If
        
        TBProduto.Update
        
        ProcCopiarComposProdServ
        
        TBLISTA.MoveNext
    Loop
    TBProduto.Close
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarServico()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_carteira where cotacao = " & IDAntigo & " and tipo = 'S' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from vendas_carteira", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBLISTA.EOF = False
        TBProduto.AddNew
        TBProduto!Tem_ordem = False
        TBProduto!Cotacao = txtId.Text
        If IsNull(TBLISTA!Desenho) = False Then TBProduto!Desenho = TBLISTA!Desenho
        If IsNull(TBLISTA!N_referencia) = False Then TBProduto!N_referencia = TBLISTA!N_referencia
        If IsNull(TBLISTA!ID_CFOP) = False Then TBProduto!ID_CFOP = TBLISTA!ID_CFOP
        If IsNull(TBLISTA!Rev_codinterno) = False Then TBProduto!Rev_codinterno = TBLISTA!Rev_codinterno
        If IsNull(TBLISTA!descricao_tecnica) = False Then TBProduto!descricao_tecnica = TBLISTA!descricao_tecnica
        If IsNull(TBLISTA!Descricao) = False Then TBProduto!Descricao = TBLISTA!Descricao
        TBProduto!Servico_cliente = TBLISTA!Servico_cliente
        If IsNull(TBLISTA!quantidade) = False Then TBProduto!quantidade = TBLISTA!quantidade
        If IsNull(TBLISTA!Desconto) = False Then TBProduto!Desconto = TBLISTA!Desconto
        If IsNull(TBLISTA!ValorDesconto) = False Then TBProduto!ValorDesconto = TBLISTA!ValorDesconto
        If IsNull(TBLISTA!preco_unitario_desconto) = False Then TBProduto!preco_unitario_desconto = TBLISTA!preco_unitario_desconto
        If IsNull(TBLISTA!preco_unitario) = False Then TBProduto!preco_unitario = TBLISTA!preco_unitario
        If IsNull(TBLISTA!preco_lote) = False Then TBProduto!preco_lote = TBLISTA!preco_lote
        TBProduto!Cidade = TBLISTA!Cidade
        If IsNull(TBLISTA!Unidade) = False Then TBProduto!Unidade = TBLISTA!Unidade
        If IsNull(TBLISTA!Unidade_com) = False Then TBProduto!Unidade_com = TBLISTA!Unidade_com
        If IsNull(TBLISTA!Familia) = False Then TBProduto!Familia = TBLISTA!Familia
        If IsNull(TBLISTA!PrazoFinal) = False Then TBProduto!PrazoFinal = TBLISTA!PrazoFinal
        If IsNull(TBLISTA!Prazo_original) = False Then TBProduto!Prazo_original = TBLISTA!Prazo_original
        If IsNull(TBLISTA!prazofinaldias) = False Then TBProduto!prazofinaldias = TBLISTA!prazofinaldias
        If IsNull(TBLISTA!Qtde_produzir) = False Then TBProduto!Qtde_produzir = TBLISTA!Qtde_produzir
        
        'Impostos
        If IsNull(TBLISTA!PIS_Serv) = False Then TBProduto!PIS_Serv = TBLISTA!PIS_Serv
        If IsNull(TBLISTA!Total_PIS_serv) = False Then TBProduto!Total_PIS_serv = TBLISTA!Total_PIS_serv
        If IsNull(TBLISTA!Cofins_Serv) = False Then TBProduto!Cofins_Serv = TBLISTA!Cofins_Serv
        If IsNull(TBLISTA!Total_Cofins_serv) = False Then TBProduto!Total_Cofins_serv = TBLISTA!Total_Cofins_serv
        If IsNull(TBLISTA!CSLL_Serv) = False Then TBProduto!CSLL_Serv = TBLISTA!CSLL_Serv
        If IsNull(TBLISTA!Total_CSLL_serv) = False Then TBProduto!Total_CSLL_serv = TBLISTA!Total_CSLL_serv
        If IsNull(TBLISTA!ISS) = False Then TBProduto!ISS = TBLISTA!ISS
        If IsNull(TBLISTA!VlrISS) = False Then TBProduto!VlrISS = TBLISTA!VlrISS
        If IsNull(TBLISTA!INSS_Serv) = False Then TBProduto!INSS_Serv = TBLISTA!INSS_Serv
        If IsNull(TBLISTA!Total_INSS_serv) = False Then TBProduto!Total_INSS_serv = TBLISTA!Total_INSS_serv
        If IsNull(TBLISTA!IRPJ_Serv) = False Then TBProduto!IRPJ_Serv = TBLISTA!IRPJ_Serv
        If IsNull(TBLISTA!Total_IRPJ_serv) = False Then TBProduto!Total_IRPJ_serv = TBLISTA!Total_IRPJ_serv
        If IsNull(TBLISTA!IRRF_Serv) = False Then TBProduto!IRRF_Serv = TBLISTA!IRRF_Serv
        If IsNull(TBLISTA!Total_IRRF_serv) = False Then TBProduto!Total_IRRF_serv = TBLISTA!Total_IRRF_serv
        If IsNull(TBLISTA!cpp) = False Then TBProduto!cpp = TBLISTA!cpp
        If IsNull(TBLISTA!Total_CPP) = False Then TBProduto!Total_CPP = TBLISTA!Total_CPP
                
        TBProduto!Tipo = "S"
        If Vendas_Proposta = True Then
            TBProduto!Liberacao = "ABERTA EM ANALISE"
        Else
            TBProduto!Liberacao = "VENDIDA"
            TBProduto!Datavendas = Date
        End If
        TBProduto!Observacoes = TBLISTA!Observacoes
        TBProduto!Obs_faturamento = TBLISTA!Obs_faturamento
        TBProduto!Antecipacao_fat = TBLISTA!Antecipacao_fat
        TBProduto!Faturamento_parcial = TBLISTA!Faturamento_parcial
        
        If Revisar = True Then
            TBLISTA!Liberacao = "REVISADA"
            TBLISTA.Update
        End If
        TBProduto.Update
        
        ProcCopiarComposProdServ
        
        TBLISTA.MoveNext
    Loop
    TBProduto.Close
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarComposProdServ()
On Error GoTo tratar_erro

Set TBComponente = CreateObject("adodb.recordset")
TBComponente.Open "Select * from vendas_carteira_composicao where ID_carteira = " & TBLISTA!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
If TBComponente.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from vendas_carteira_composicao", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBComponente.EOF = False
        TBGravar.AddNew
        TBGravar!ID_carteira = TBProduto!CODIGO
        TBGravar!Codigo_interno = TBComponente!Codigo_interno
        TBGravar!Descricao = TBComponente!Descricao
        TBGravar!Un = TBComponente!Un
        TBGravar!valor_unitario = TBComponente!valor_unitario
        TBGravar!quantidade = TBComponente!quantidade
        TBGravar!Valor_total = TBComponente!Valor_total
        TBGravar.Update
        TBComponente.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarTotais()
On Error GoTo tratar_erro

TBCotacao!dbl_Valor_Total_Produtos = TBProposta!dbl_Valor_Total_Produtos
TBCotacao!dbl_valor_total_servicos = TBProposta!dbl_valor_total_servicos
TBCotacao!TotalDesconto = TBProposta!TotalDesconto
TBCotacao!SubTotal = TBProposta!SubTotal
TBCotacao!dbl_valor_total = TBProposta!dbl_valor_total

'Impostos produtos
TBCotacao!dbl_Base_ICMS = TBProposta!dbl_Base_ICMS
TBCotacao!dbl_Valor_ICMS = TBProposta!dbl_Valor_ICMS
TBCotacao!dbl_Valor_Total_IPI = TBProposta!dbl_Valor_Total_IPI
TBCotacao!Total_PIS_prod = TBProposta!Total_PIS_prod
TBCotacao!Total_Cofins_prod = TBProposta!Total_Cofins_prod
TBCotacao!Total_CSLL_prod = TBProposta!Total_CSLL_prod

'Impostos serviços
TBCotacao!Total_PIS_serv = TBProposta!Total_PIS_serv
TBCotacao!Total_Cofins_serv = TBProposta!Total_Cofins_serv
TBCotacao!Total_CSLL_serv = TBProposta!Total_CSLL_serv
TBCotacao!VlrTotaliss = TBProposta!VlrTotaliss
TBCotacao!Total_INSS_serv = TBProposta!Total_INSS_serv
TBCotacao!Total_IRPJ_serv = TBProposta!Total_IRPJ_serv
TBCotacao!Total_IRRF_serv = TBProposta!Total_IRRF_serv

'Impostos faturamento
TBCotacao!Total_DAS = TBProposta!Total_DAS

'Retenção de PIS/Cofins
TBCotacao!Total_retencao_PIS = IIf(IsNull(TBProposta!Total_retencao_PIS), 0, TBProposta!Total_retencao_PIS)
TBCotacao!Total_retencao_Cofins = IIf(IsNull(TBProposta!Total_retencao_Cofins), 0, TBProposta!Total_retencao_Cofins)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtCotacao = "" Then
    USMsgBox ("Informe a proposta comercial antes de revisar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Vendas_proposta", "Cotacao = " & txtId, "mesma", "a proposta", "revisar", False, False) = False Then Exit Sub
If txtStatus <> "ABERTA EM ANALISE" Then
    USMsgBox ("Só é permitido revisar proposta com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente criar uma revisão da proposta " & txtCotacao.Text & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Revisar = True
    IDAntigo = txtId.Text
    Set TBProposta = CreateObject("adodb.recordset")
    TBProposta.Open "Select * from vendas_proposta where cotacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBProposta.EOF = False Then
        Set TBCotacao = CreateObject("adodb.recordset")
        TBCotacao.Open "Select * from vendas_proposta", Conexao, adOpenKeyset, adLockOptimistic
        txtrevisao.Text = TBProposta!Revisao + 1
        TBCotacao.AddNew
        TBCotacao!Ncotacao = txtCotacao.Text
        TBCotacao!ordenarproposta = TBProposta!ordenarproposta
        TBCotacao!Data = Date
        TBCotacao!Revisao = txtrevisao.Text
        TBCotacao!Responsavel = pubUsuario
        ProcCopiarInfGerais
        ProcCopiarTotais
        TBCotacao.Update
        txtId.Text = TBCotacao!Cotacao
        ProcCopiarDetalhes
        ProcCopiarItem
        ProcCopiarServico
        TBCotacao.Close
        
        TBProposta!status = "REVISADA"
        TBProposta!dataalteracao = Date
        TBProposta.Update
    End If
    TBProposta.Close
    
    USMsgBox ("Proposta revisada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Revisar"
    ID_documento = txtId
    Documento = "Nº proposta: " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_proposta where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then ProcPuxaDados
    TBAbrir.Close
    ProcAtualizalistaProdutos (1)
    ProcAtualizalistaServicos (1)
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmdvendedor_Click()
On Error GoTo tratar_erro

VE = True
VI = False
'Vendas_PI = False
frmVendas_Lista_Vendedores.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdVendedor_Interno_Click()
On Error GoTo tratar_erro

VI = True
VE = False
'Vendas_PI = False
frmVendas_Lista_Vendedores.Show 1

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
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiar
            Case vbKeyF8: If Vendas_Proposta = True Then ProcRevisao
            Case vbKeyF9: If Vendas_Proposta = True Then ProcEmitirPI
            Case vbKeyF10: ProcCancelaPI
            Case vbKeyF11: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF12: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, IIf(Vendas_Proposta = True, "Vendas/Proposta comercial", "Vendas/Pedido interno")
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: ProcSalvarComercial
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcFinanceiro
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_prod
            Case vbKeyF2:
                If Frame1(10).Enabled = True Then cmdlistaproduto_Click
            Case vbKeyF3: ProcSalvar_prod
            Case vbKeyF4: If cmbOpcao_lista_prod = "Excluir" Then ProcExcluir_prod
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcEstrutura
            Case vbKeyF8: If Vendas_Proposta = True Then ProcEmitirPI_prod
            Case vbKeyF9: ProcCancelaPI_prod
            Case vbKeyF10: ProcComposicao_prod
            Case vbKeyF11: If cmbOpcao_lista_prod = "Status" Then ProcStatus
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_serv
            Case vbKeyF2:
                If Frame1(11).Enabled = True Then cmdlistaservicos_Click
            Case vbKeyF3: ProcSalvar_serv
            Case vbKeyF4: If cmbOpcao_lista_serv = "Excluir" Then ProcExcluir_Serv
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Vendas_Proposta = True Then ProcEmitirPI_serv
            Case vbKeyF8: ProcCancelaPI_serv
            Case vbKeyF9: ProcComposicao_serv
            Case vbKeyF11: If cmbOpcao_lista_serv = "Status" Then ProcStatus
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
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

Private Sub Form_Load()
On Error GoTo tratar_erro

LiberarAlteracao = False
ProcCarregaToolBar1 Me, 15195, 19, True
ProcCarregaToolBar2 Me, 15195, 9, True
ProcCarregaToolBar3 Me, 15195, 16, True
ProcCarregaToolBar4 Me, 15195, 15, True
ProcCarregaToolBar5 Me, 15195, 10, True
ProcOrganizaFormPI_Proposta
ProcCarregaComboEmpresa Cmb_empresa, False
Direitos
SSTab1.Tab = 0
SSTab2.Tab = 0
ProcLimpaVariaveisPrincipais

Cmb_opcao_lista = "Validação"
cmbOpcao_lista_prod = "Excluir"
cmbOpcao_lista_serv = "Excluir"

ProcCarregaComboMoeda
ProcCarregaComboProduto
ProcCarregaComboServico
ProcCarregaCamposCombo
ProcCarregaComboUF txtuf, "UF is not null", ""

ProcRemoveObjetosResize Me
Frame1(0).Visible = False
Frame1(2).Visible = False

'========================================
' Verifica se controla venda ao cliente
'========================================
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * from empresa where Empresa = '" & Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAcessos.EOF = False Then
ClienteVendedor = IIf(IsNull(TBAcessos!ClienteVendedor), 0, TBAcessos!ClienteVendedor)
SemEstoque = IIf(IsNull(TBAcessos!ClienteVendedor), 0, TBAcessos!ClienteVendedor)
End If
TBAcessos.Close

cmdVendedor_Interno.Visible = IIf(ClienteVendedor = True, False, True)
txtComissao.Locked = IIf(ClienteVendedor = True, False, True)

    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
cmbfamiliaservico.ListIndex = -1
Cmb_cidade_servico.ListIndex = -1
txtunservico.ListIndex = -1
Cmb_un_com_serv.ListIndex = -1
If txtId <> 0 Then
    Set TBOSC = CreateObject("adodb.recordset")
    TBOSC.Open "Select Transportadora, Tipo_transp FROM vendas_comercial WHERE cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBOSC.EOF = False Then
        If IsNull(TBOSC!Transportadora) = False And TBOSC!Transportadora <> "" Then
            Select Case TBOSC!Tipo_transp
                Case "C": Cmb_tipo_transp = "Cliente"
                Case "F": Cmb_tipo_transp = "Fornecedor"
                Case "E": Cmb_tipo_transp = "Empresa"
            End Select
        End If
    End If
    TBOSC.Close
End If
If txtid_produto <> 0 Then
    Set TBOSC = CreateObject("adodb.recordset")
    TBOSC.Open "Select Familia, Unidade, Unidade_com from vendas_carteira where Codigo = " & txtid_produto, Conexao, adOpenKeyset, adLockOptimistic
    If TBOSC.EOF = False Then
        If IsNull(TBOSC!Familia) = False And TBOSC!Familia <> "" Then cmbfamilia.Text = TBOSC!Familia
        If IsNull(TBOSC!Unidade) = False And TBOSC!Unidade <> "" Then cmbun.Text = TBOSC!Unidade
        If IsNull(TBOSC!Unidade_com) = False And TBOSC!Unidade_com <> "" Then Cmb_un_com.Text = TBOSC!Unidade_com
    End If
    TBOSC.Close
End If
If txtid_servico <> 0 Then
    Set TBOSC = CreateObject("adodb.recordset")
    TBOSC.Open "Select Familia, Cidade, Unidade, Unidade_com from vendas_carteira where Codigo = " & txtid_servico, Conexao, adOpenKeyset, adLockOptimistic
    If TBOSC.EOF = False Then
        If IsNull(TBOSC!Familia) = False And TBOSC!Familia <> "" Then cmbfamiliaservico.Text = TBOSC!Familia
        If IsNull(TBOSC!Cidade) = False And TBOSC!Cidade <> "" Then Cmb_cidade_servico = TBOSC!Cidade
        If IsNull(TBOSC!Unidade) = False And TBOSC!Unidade <> "" Then txtunservico.Text = TBOSC!Unidade
        If IsNull(TBOSC!Unidade_com) = False And TBOSC!Unidade_com <> "" Then Cmb_un_com_serv.Text = TBOSC!Unidade_com
    End If
    TBOSC.Close
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

ProcLimpar
txtId.Text = TBAbrir!Cotacao
If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
txtStatus.Text = IIf(IsNull(TBAbrir!status), "", TBAbrir!status)
Caption = "Vendas - " & IIf(Vendas_Proposta = True, "Proposta comercial - (Proposta : ", "Pedido interno - (Pedido interno : ") & TBAbrir!Ncotacao & " - Rev. : " & TBAbrir!Revisao & ")"
txtCotacao.Text = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
txtrevisao.Text = IIf(IsNull(TBAbrir!Revisao), "", TBAbrir!Revisao)
If Vendas_Proposta = True Then
    txtDatavendas = IIf(IsNull(TBAbrir!Datavendas), "", Format(TBAbrir!Datavendas, "dd/mm/yyyy"))
Else
    txtDatavendas_PI = IIf(IsNull(TBAbrir!Datavendas), Date, Format(TBAbrir!Datavendas, "dd/mm/yyyy"))
End If

RegimeEmpresa_PI = IIf(IsNull(TBAbrir!Regime), 0, TBAbrir!Regime)
txtIDcliente.Text = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
txtCliente.Text = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)

NomeCampo = "o estado"
If IsNull(TBAbrir!UF) = False And TBAbrir!UF <> "" Then
    If TBAbrir!UF <> txtuf Then
        txtuf.Text = TBAbrir!UF
        
        If Vendas_Proposta = True Then
            NomeCampo = "a Cidade"
            If TBAbrir!UF <> "EX" Then
                cmbCidade.Visible = True
                txtCidade.Visible = False
                If IsNull(TBAbrir!Cidade) = False And TBAbrir!Cidade <> "" Then cmbCidade = TBAbrir!Cidade
            Else
                cmbCidade.Visible = False
                txtCidade.Visible = True
                txtCidade.Text = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
            End If
        Else
            txtCidade.Text = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
        End If
    End If
Else
    txtuf.ListIndex = -1
End If

1:
    txtRemetente.Text = IIf(IsNull(TBAbrir!Remetente), "", TBAbrir!Remetente)
    txtFax.Text = IIf(IsNull(TBAbrir!Fax), "", TBAbrir!Fax)
    txtEmail.Text = IIf(IsNull(TBAbrir!Email), "", TBAbrir!Email)
    txtRef.Text = IIf(IsNull(TBAbrir!Referente), "", TBAbrir!Referente)
    
    If IsNull(TBAbrir!Tipo_endereco) = False And TBAbrir!Tipo_endereco <> "" Then
    cmbTipo_endereco = TBAbrir!Tipo_endereco
    End If
    
    txtendereco.Text = IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
    txtNumero.Text = IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero)
    txtComplemento.Text = IIf(IsNull(TBAbrir!complemento), "", TBAbrir!complemento)
    If IsNull(TBAbrir!Tipo_bairro) = False And TBAbrir!Tipo_bairro <> "" Then cmbTipo_bairro = TBAbrir!Tipo_bairro
    txtBairro.Text = IIf(IsNull(TBAbrir!Bairro), "", TBAbrir!Bairro)
    txttelefone.Text = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)
    txtVI.Text = IIf(IsNull(TBAbrir!VI), "", TBAbrir!VI)
    txtVE.Text = IIf(IsNull(TBAbrir!VE), "", TBAbrir!VE)
    txtregiao.Text = IIf(IsNull(TBAbrir!regiao), "", TBAbrir!regiao)
    txtdepartamento.Text = IIf(IsNull(TBAbrir!Departamento), "", TBAbrir!Departamento)
    
    If IsNull(TBAbrir!Tipo_cliente) = False And TBAbrir!Tipo_cliente <> "" Then
    txttipocliente.Text = TBAbrir!Tipo_cliente
    End If
    
    txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    txtVend_Ext.Text = IIf(IsNull(TBAbrir!Vend_ext), "", TBAbrir!Vend_ext)
    txtvend_Int.Text = IIf(IsNull(TBAbrir!vend_int), "", TBAbrir!vend_int)
    txt_observacoes.Text = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
    txtreferente.Text = IIf(IsNull(TBAbrir!Ref), "", TBAbrir!Ref)
    txt_dataelaborado.Value = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yyyy"))
    txt_datamodificado.Text = IIf(IsNull(TBAbrir!dataalteracao), "", Format(TBAbrir!dataalteracao, "dd/mm/yy"))
    txtcnpj.Text = IIf(IsNull(TBAbrir!CNPJ_CPF), "", TBAbrir!CNPJ_CPF)
    txtCEP.Text = IIf(IsNull(TBAbrir!CEP), "", TBAbrir!CEP)
    txtIE.Text = IIf(IsNull(TBAbrir!RG_IE), "", TBAbrir!RG_IE)
    
    
    If Vendas_Proposta = True Then
        txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
        txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
    Else
        txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacaoPI), "", TBAbrir!DtValidacaoPI)
        txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacaoPI), "", TBAbrir!RespValidacaoPI)
    End If
    TabelaSN_PI = IIf(IsNull(TBAbrir!TabelaSN), 0, TBAbrir!TabelaSN)
    
    With Label1(32)
        If TBAbrir!status <> "CANCELADA" And TBAbrir!status <> "PERDIDA P/ PRAZO" And TBAbrir!status <> "PERDIDA P/ PREÇO" Then
            .Caption = "Revisada em"
        Else
            Select Case TBAbrir!status
                Case "CANCELADA": .Caption = "Cancelada em"
                Case "PERDIDA P/ PRAZO": .Caption = "Perdida p/ prazo em"
                Case "PERDIDA P/ PREÇO": .Caption = "Perdida p/ preço em"
            End Select
        End If
        Label1(32).Left = txt_datamodificado.Left + (txt_datamodificado.Width / 2) - (Label1(32).Width / 2)
    End With
    
    Frame1(1).Enabled = True
    Novo_PI = False
    
Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        If NomeCampo = "o estado" Then
            USMsgBox ("Este cliente não é compativel com esta CFOP, favor revisar."), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Não foi encontrado " & NomeCampo & " desse cliente."), vbExclamation, "CAPRIND v5.0"
        End If
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBCotacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
If LiberarAlteracao = False Then
If txtStatus = "FATURADA" Or txtStatus = "FATURADA PARCIAL" Then Exit Sub


If Vendas_Proposta = True Then If txtStatus = "CANCELADA" Or txtStatus = "PERDIDA P/ PRAZO" Or txtStatus = "PERDIDA P/ PREÇO" Then TBCotacao!dataalteracao = Date Else TBCotacao!dataalteracao = Null
End If

TBCotacao!status = txtStatus.Text
TBCotacao!Revisao = txtrevisao.Text
'TBCotacao!VI = txtVI
TBCotacao!VE = txtVE
TBCotacao!regiao = txtregiao.Text

TBCotacao!IDCliente = IIf(txtIDcliente = "", 0, txtIDcliente)

TBCotacao!Cliente = Replace(txtCliente.Text, "'", " ")
TBCotacao!Remetente = IIf(txtRemetente = "", Null, txtRemetente)
TBCotacao!Referente = IIf(txtRef = "", Null, txtRef)
TBCotacao!Fax = txtFax.Text
TBCotacao!Email = IIf(txtEmail.Text = "", Null, LCase(txtEmail.Text))

TBCotacao!Tipo_endereco = IIf(cmbTipo_endereco = "", Null, cmbTipo_endereco)
TBCotacao!Endereco = IIf(txtendereco = "", Null, txtendereco)
TBCotacao!Numero = IIf(txtNumero = "", Null, txtNumero)
TBCotacao!complemento = IIf(txtComplemento = "", Null, txtComplemento)
TBCotacao!Tipo_bairro = IIf(cmbTipo_bairro = "", Null, cmbTipo_bairro)
TBCotacao!Bairro = IIf(txtBairro = "", Null, txtBairro)
If txtCidade.Visible = True Then TBCotacao!Cidade = IIf(txtCidade = "", Null, txtCidade) Else TBCotacao!Cidade = IIf(cmbCidade = "", Null, cmbCidade)
TBCotacao!telefone = txttelefone.Text
TBCotacao!Departamento = txtdepartamento.Text
TBCotacao!UF = IIf(txtuf = "", Null, txtuf)
TBCotacao!Tipo_cliente = txttipocliente.Text
TBCotacao!Datavendas = IIf(Vendas_PI = True, txtDatavendas_PI, IIf(txtDatavendas = "", Null, txtDatavendas))
TBCotacao!Obs = IIf(txt_observacoes = "", Null, txt_observacoes)
TBCotacao!Ref = IIf(txtreferente = "", Null, txtreferente)
TBCotacao!Vend_ext = txtVend_Ext.Text
TBCotacao!vend_int = txtvend_Int.Text
TBCotacao!Data = IIf(txt_dataelaborado = "", Date, txt_dataelaborado)
TBCotacao!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
'If IsNull(TBCotacao!ordenarproposta) = True Or TBCotacao!ordenarproposta = 0 Then TBCotacao!ordenarproposta = txtID.Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro
  
frmVendas_PI_lista.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarContato()
On Error GoTo tratar_erro

USMsgBox "Escolha seu contato com o cliente", vbInformation, "CAPRIND v5.0"

If txtIDcliente.Text <> "" And txtIDcliente.Text <> "0" Then
    Analise_critica = False
    Telemarketing = False
    Qualidade_PPAP_PSW = False
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    frmVendas_propostaII_contato.Show 1
End If

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
If FunVerifRegimeEmpresa(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        USMsgBox ("Não é permitido criar " & IIf(Vendas_Proposta = True, "proposta", "pedido") & ", pois não existe nenhuma tabela do simples nacional ativa."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If
ProcLimpar
ProcLimparTudo
Frame1(1).Enabled = True
Novo_PI = True
continuar = True
txt_dataelaborado.Value = Date
txtResponsavel = pubUsuario
If Vendas_Proposta = True Then txtStatus.Text = "ABERTA EM ANALISE" Else txtStatus.Text = "VENDIDA"
txtrevisao = 0
'============================================
' Localizar o cliente
'============================================
If continuar = True Then
ProcLocalizarCliente
End If
'============================================
' Escolher o contato com cliente
'============================================
If continuar = True Then
ProcLocalizarContato
End If
'============================================
' Localizar o vendedor interno
'============================================
If continuar = True Then
ProcLocalizaVendedorInterno
End If
'============================================
' Localizar o vendedor externo
'============================================
If continuar = True Then
ProcLocalizarVendedorExterno
End If
'============================================
' Salvar a proposta comercial
'============================================
If continuar = True Then
ProcSalvar
End If
'============================================

continuar = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizaVendedorInterno2()
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
'TBClientes.Open "Select * from usuarios where Usuario = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockOptimistic
TBClientes.Open "Select * from Vendas_Vendedores_Clientes where IDCliente = " & txtIDcliente & "", Conexao, adOpenKeyset, adLockOptimistic

If TBClientes.EOF = False Then
txtComissao.Text = TBClientes!Comissao
Set TBAcessos = CreateObject("adodb.recordset")
'TBAcessos.Open "Select * from Vendas_Vendedores where Vendedor = '" & TBClientes!Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
TBAcessos.Open "Select * from Vendas_Vendedores where ID = " & TBClientes!IDvendedor & "", Conexao, adOpenKeyset, adLockOptimistic

If TBAcessos.EOF = False Then
txtVI.Text = TBAcessos!ID
txtvend_Int.Text = TBAcessos!vendedor
Else
USMsgBox "Escolha o vendedor interno", vbInformation, "CAPRIND v5.0"
cmdVendedor_Interno_Click
End If
TBAcessos.Close
End If
TBClientes.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarVendedorExterno()
On Error GoTo tratar_erro

USMsgBox "Escolha seu vendedor externo", vbInformation, "CAPRIND v5.0"

VE = True
VI = False
'Vendas_PI = False
frmVendas_Lista_Vendedores.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcLocalizarCliente()
On Error GoTo tratar_erro

Sit_REG = 1
If Vendas_PI = True Then
    ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False
Else
    ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False
End If

'================================================================
' Se não amarra cliente ao vendedor interno
'================================================================
If ClienteVendedor = False Then
 frmVendas_LocalizarCliente.Show 1
 If txtIDcliente.Text = "" Then
  USMsgBox "É obrigatório escolher um cliente para o pedido interno", vbCritical, "CAPRIND v5.0"
  continuar = False
  Exit Sub
 End If
End If

'================================================================
' Se amarra cliente ao vendedor interno
'================================================================
If ClienteVendedor = True Then
 frmVendas_Vendedores_LocalizarCliente.Show 1
 If txtIDcliente.Text = "" Then
  USMsgBox "É obrigatório escolher um cliente para o pedido interno", vbCritical, "CAPRIND v5.0"
  continuar = False
  Exit Sub
 End If
End If

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
If Vendas_Proposta = True Then
    If txtStatus = "VENDIDA" Then
        USMsgBox ("Não é permitida a alteração de proposta vendida."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TextoPadrao = "a proposta esta "
Else
    TextoPadrao = "o pedido interno esta "
End If
If txtStatus.Text = "REVISADA" Or txtStatus.Text = "FATURADA" Or txtStatus = "FATURADA PARCIAL" And LiberarAlteracao = False Then
    USMsgBox ("Não é permitido alterar, pois " & TextoPadrao & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1(1).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtStatus.Text = "" Then
    NomeCampo = "o status"
    ProcVerificaAcao
    cmdstatus.SetFocus
    Exit Sub
End If
If Vendas_Proposta = True Then
    'proposta cliente sem cadastro
    If txtCliente = "" Then
        NomeCampo = "o cliente"
        ProcVerificaAcao
        cmdAdicionarCliente_Click
        Exit Sub
    End If
Else
    'pedido aceita apenas cliente com cadastro
    If txtIDcliente = "" Or txtIDcliente = "0" Then
        NomeCampo = "o cliente"
        ProcVerificaAcao
        cmdAdicionarCliente_Click
        Exit Sub
    End If
    
    If txtCidade <> "" And txtuf <> "" And txtuf <> "EX" Then
        If FunVerificaCidade(txtCidade, txtuf) = False Then Exit Sub
    End If
End If
If txttipocliente = "" Then
    NomeCampo = "o tipo do cliente"
    ProcVerificaAcao
    txttipocliente.SetFocus
    Exit Sub
End If
If txtvend_Int.Text = "" Then
    NomeCampo = "o vendedor interno"
    ProcVerificaAcao
    cmdVendedor_Interno_Click
    Exit Sub
End If
If txtVend_Ext.Text = "" Then
    NomeCampo = "o vendedor externo"
    ProcVerificaAcao
    Cmdvendedor_Click
    Exit Sub
End If
If txtregiao.Text = "" Then
    NomeCampo = "a região do vendedor externo"
    ProcVerificaAcao
    Cmdvendedor_Click
    Exit Sub
End If

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * FROM Clientes WHERE idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Left(Tipo, 1) = 'J' and idTipoEmpresa = 1", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    If FunVerifRegimeTribCliForn(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
End If
TBClientes.Close

NumeroCotacao = txtCotacao
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM vendas_proposta where Cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = True Then
    RegimeEmpresa_PI = FunVerifRegimeEmpresa(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
    If RegimeEmpresa_PI = 1 Then
        'Verifica se existe mais de uma tabela do simples cadastrada
        Contador = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                TabelaSN_PI = TBAbrir!Tabela
                Contador = Contador + 1
                TBAbrir.MoveNext
            Loop
            If Contador > 1 Then
                USMsgBox ("Favor informar a tabela do simples nacional utilizada para " & IIf(Vendas_Proposta = True, "essa proposta", "esse pedido") & "."), vbInformation, "CAPRIND v5.0"
                Vendas_Programacao = False
                frmVendas_proposta_tabelaSN.Show 1
            End If
        End If
        TBAbrir.Close
    End If
        
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Ncotacao from vendas_proposta where Year(Data) = '" & Year(Date) & "' order by Ordenarproposta desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Cotacao = Left(TBAbrir!Ncotacao, Len(TBAbrir!Ncotacao) - 3) + 1
    Else
        Cotacao = 1
    End If
    Ano = Right(Year(Date), 2)
    Select Case Len(Cotacao)
        Case 1: NumeroCotacao = "000" & Cotacao & "/" & Ano
        Case 2: NumeroCotacao = "00" & Cotacao & "/" & Ano
        Case 3: NumeroCotacao = "0" & Cotacao & "/" & Ano
        Case 4: NumeroCotacao = Cotacao & "/" & Ano
        Case 5: NumeroCotacao = Cotacao & "/" & Ano
    End Select
    txtCotacao = NumeroCotacao
    TBCotacao.AddNew
    TBCotacao!Ncotacao = txtCotacao
    TBCotacao!Tipo = IIf(Vendas_Proposta = True, "PR", "PE")
    TBCotacao!TabelaSN = TabelaSN_PI
    TBCotacao!Regime = RegimeEmpresa_PI
Else
    If LiberarAlteracao = False Then
    If Vendas_Proposta = True Then
        If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesma", "a proposta", False) = False Then Exit Sub
    Else
        If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "o pedido interno", True) = False Then Exit Sub
    End If
    End If
    
    'Verifica se é um pedido revisado e não deixa alterar os dados principais
    If txtrevisao > 0 And TBCotacao!IDCliente <> 0 And TBCotacao!IDCliente <> txtIDcliente And LiberarAlteracao = False Then
        USMsgBox ("Não é permitido alterar o cliente " & IIf(Vendas_Proposta = True, "desta proposta, pois a mesma", "deste pedido, pois o mesmo") & " é uma revisão."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    If TBCotacao!IDCliente <> txtIDcliente Then
        Conexao.Execute "UPDATE Vendas_Comercial set Local_entrega = NULL, Local_cobranca = NULL, ID_entrega = 0, ID_cobranca = 0 WHERE cotacao = " & TBCotacao!Cotacao
        ProcSalvarDadosComerciaisCliente
    End If
End If
ProcEnviaDados
If txtStatus.Text = "PORTAL ELETRONICO" Then
    ValorBox = InputBox("Informe o valor total " & IIf(Vendas_Proposta = True, "da proposta.", "do pedido."))
    If ValorBox <> "" Then TBCotacao!dbl_Valor_Total_Produtos = ValorBox
End If
TBCotacao.Update
txtId.Text = TBCotacao!Cotacao
If Novo_PI = True Then
    Conexao.Execute "Update Vendas_proposta set ordenarproposta = " & TBCotacao!Cotacao & " where cotacao = " & TBCotacao!Cotacao
    If txtIDcliente <> "" Then ProcSalvarDadosComerciaisCliente
End If

txtCotacao = NumeroCotacao
Caption = "Vendas - " & IIf(Vendas_PI = True, "Pedido interno - (Pedido interno : ", "Proposta comercial - (Proposta : ") & txtCotacao & " - Rev. : " & txtrevisao & ")"
TBCotacao.Close

ProcgravarItem

If Novo_PI = True Then
    USMsgBox IIf(Vendas_PI = True, "Novo pedido interno cadastrado", "Nova proposta cadastrada") & " com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_PI_Localizar = "Select * FROM vendas_proposta where Cotacao = " & txtId.Text
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista2 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista2)
        Lista.SetFocus
    End If
End If
1:
    '==================================
    Modulo = Formulario
    ID_documento = txtId
    Documento = IIf(Vendas_PI = True, "Nº pedido", "Nº proposta") & ": " & txtCotacao & " - Rev.: " & txtrevisao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_PI = False
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_proposta where cotacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If TBAbrir!dataalteracao <> "" Then txt_datamodificado.Text = TBAbrir!dataalteracao Else txt_datamodificado.Text = ""
        If TBAbrir!status = "CANCELADA" Then Label1(32).Caption = "Cancelada em"
        If TBAbrir!status = "PERDIDA P/ PRAZO" Then Label1(32).Caption = "Perdida p/ prazo em"
        If TBAbrir!status = "PERDIDA P/ PREÇO" Then Label1(32).Caption = "Perdida p/ preço em"
        If TBAbrir!status <> "CANCELADA" And TBAbrir!status <> "PERDIDA P/ PRAZO" And TBAbrir!status <> "PERDIDA P/ PREÇO" Then Label1(32).Caption = "Revisada em"
    End If
    TBAbrir.Close

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarDadosComerciaisCliente()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from Clientes_DadosComerciais where idcliente = " & txtIDcliente & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from vendas_comercial where cotacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        If TBGravar.EOF = True Then TBGravar.AddNew
        ProcEnviaDadosComerciaisPadrao
        TBGravar!Cotacao = txtId
        
        Set TBVendas = CreateObject("adodb.recordset")
        TBVendas.Open "Select Tipo_transp, idTransp, txt_transportadora from clientes where idcliente = " & txtIDcliente & " and txt_transportadora is not null", Conexao, adOpenKeyset, adLockOptimistic
        If TBVendas.EOF = False Then
            TBGravar!Tipo_transp = IIf(IsNull(TBVendas!Tipo_transp), "", TBVendas!Tipo_transp)
            TBGravar!IdIntTransp = IIf(IsNull(TBVendas!idTransp), 0, TBVendas!idTransp)
            TBGravar!Transportadora = IIf(IsNull(TBVendas!txt_transportadora), "", TBVendas!txt_transportadora)
        End If
        TBVendas.Close
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from clientes_entrega where idcliente = " & txtIDcliente & " and Tipo = 'C' order by identrega", Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            TBGravar!ID_entrega = TBCiclo!identrega
            If IsNull(TBCiclo!Tipo_endereco) = False And TBCiclo!Tipo_endereco <> "" Then
                Endereco = TBCiclo!Tipo_endereco & ": " & IIf(IsNull(TBCiclo!endereco_entrega), "", TBCiclo!endereco_entrega)
            Else
                Endereco = IIf(IsNull(TBCiclo!endereco_entrega), "", TBCiclo!endereco_entrega)
            End If
            If IsNull(TBCiclo!Tipo_bairro) = False And TBCiclo!Tipo_bairro <> "" Then
                Bairro = TBCiclo!Tipo_bairro & ": " & IIf(IsNull(TBCiclo!bairro_entrega), "", TBCiclo!bairro_entrega)
            Else
                Bairro = IIf(IsNull(TBCiclo!bairro_entrega), "", TBCiclo!bairro_entrega)
            End If
            Endereco1 = Endereco & " - " & IIf(IsNull(TBCiclo!Numero), "", TBCiclo!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBCiclo!cidade_entrega), "", TBCiclo!cidade_entrega) & " - " & IIf(IsNull(TBCiclo!uf_entrega), "", TBCiclo!uf_entrega) & " - " & IIf(IsNull(TBCiclo!cep_entrega), "", TBCiclo!cep_entrega)
            TBGravar!Local_entrega = Endereco1
        End If
        TBCiclo.Close
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from clientes_cobranca where idcliente = " & txtIDcliente & " and Tipo = 'C' order by idcobranca", Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            TBGravar!ID_Cobranca = TBCiclo!idCobranca
            If IsNull(TBCiclo!Tipo_endereco) = False And TBCiclo!Tipo_endereco <> "" Then
                Endereco = TBCiclo!Tipo_endereco & ": " & IIf(IsNull(TBCiclo!endereco_Cobranca), "", TBCiclo!endereco_Cobranca)
            Else
                Endereco = IIf(IsNull(TBCiclo!endereco_Cobranca), "", TBCiclo!endereco_Cobranca)
            End If
            If IsNull(TBCiclo!Tipo_bairro) = False And TBCiclo!Tipo_bairro <> "" Then
                Bairro = TBCiclo!Tipo_bairro & ": " & IIf(IsNull(TBCiclo!bairro_Cobranca), "", TBCiclo!bairro_Cobranca)
            Else
                Bairro = IIf(IsNull(TBCiclo!bairro_Cobranca), "", TBCiclo!bairro_Cobranca)
            End If
            Endereco1 = Endereco & " - " & IIf(IsNull(TBCiclo!Numero), "", TBCiclo!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBCiclo!cidade_Cobranca), "", TBCiclo!cidade_Cobranca) & " - " & IIf(IsNull(TBCiclo!uf_Cobranca), "", TBCiclo!uf_Cobranca) & " - " & IIf(IsNull(TBCiclo!cep_Cobranca), "", TBCiclo!cep_Cobranca)
            TBGravar!Local_Cobranca = Endereco1
        End If
        TBCiclo.Close
        
        TBGravar.Update
    End If
    TBGravar.Close
Else
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from vendas_comercial where cotacao = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar!Cotacao = txtId
        TBGravar!Moeda = "REAL"
        TBGravar!Valor_moeda = 1
        TBGravar.Update
    End If
    TBGravar.Close
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosComerciaisPadrao()
On Error GoTo tratar_erro

TBGravar!calculos = IIf(IsNull(TBAbrir!calculos), Null, TBAbrir!calculos)
TBGravar!impostos = IIf(IsNull(TBAbrir!impostos), Null, TBAbrir!impostos)
TBGravar!condicoes = IIf(IsNull(TBAbrir!condicoes), Null, TBAbrir!condicoes)
TBGravar!garantia = IIf(IsNull(TBAbrir!garantia), Null, TBAbrir!garantia)
TBGravar!reajuste = IIf(IsNull(TBAbrir!reajuste), Null, TBAbrir!reajuste)
TBGravar!transporte = IIf(IsNull(TBAbrir!transporte), Null, TBAbrir!transporte)
TBGravar!validade = IIf(IsNull(TBAbrir!validade), Null, TBAbrir!validade)
TBGravar!Moeda = "REAL"
TBGravar!Valor_moeda = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparProdutos(Produtos2 As Boolean)
On Error GoTo tratar_erro

If Produtos2 = False Then
    IDAnalise = 0
    txtid_produto = 0
    Txt_analise = ""
    txtNomenclatura.Text = ""
    Txt_n_serie = ""
    txtpccliente = ""
    Caminho_PC_prod_PI = ""
    txtPrazo_Produto = ""
    mskprazo.Text = "__/__/____"
    If Novo_PI1 = True Then
        If Chk_PC_prod.Value = 1 Then
            Set TBOSC = CreateObject("adodb.recordset")
            TBOSC.Open "Select PCcliente, Caminho_PCCliente from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'P' and PCcliente IS NOT NULL order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBOSC.EOF = False Then
                 txtpccliente = TBOSC!PCCliente
                 Caminho_PC_prod_PI = IIf(IsNull(TBOSC!Caminho_PCCliente), "", TBOSC!Caminho_PCCliente)
            End If
            TBOSC.Close
        End If
        If Chk_prazo_prod.Value = 1 Then
            Set TBOSC = CreateObject("adodb.recordset")
            TBOSC.Open "Select Prazofinaldias, Prazofinal from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'P' and (Prazofinaldias IS NOT NULL or prazofinal IS NOT NULL) order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBOSC.EOF = False Then
                 If Vendas_PI = True Then mskprazo = TBOSC!PrazoFinal Else txtPrazo_Produto = TBOSC!prazofinaldias
            End If
            TBOSC.Close
        End If
    End If
End If
txtvalorunitario.Text = ""
txtRev_cod = ""
cmbReferencia.Clear
txtreferencia = ""
Txt_ID_CFOP_prod = ""
Txt_CFOP_prod = ""
Txt_natureza_operacao_prod = ""
Permitido = False
If Novo_PI1 = True Then
    If Chk_CFOP_prod.Value = 1 Then
        Set TBOSC = CreateObject("adodb.recordset")
        TBOSC.Open "Select ID_CFOP from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'P' and ID_CFOP IS NOT NULL order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBOSC.EOF = False Then
            Txt_ID_CFOP_prod = TBOSC!ID_CFOP
            Permitido = True
        End If
        TBOSC.Close
    End If
    If Permitido = False Then
        'Verifica CFOP vinculada ao cliente
        Set TBOSC = CreateObject("adodb.recordset")
        TBOSC.Open "Select IDCFOP FROM Clientes_DadosComerciais where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBOSC.EOF = False Then
            Txt_ID_CFOP_prod = IIf(IsNull(TBOSC!IDCFOP), "", TBOSC!IDCFOP)
        End If
        TBOSC.Close
    End If
End If

Txt_ID_CF = ""
Txt_CF = ""
Cmb_CST_ICMS.ListIndex = -1
OPTnovo.Value = 0
OPTnovoman.Value = 0

Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select ID_CFOP, Txt_descricao, Retorno FROM tbl_NaturezaOperacao where IDCountCfop = " & IIf(Txt_ID_CFOP_prod = "", 0, Txt_ID_CFOP_prod), Conexao, adOpenKeyset, adLockOptimistic
If TBOSC.EOF = False Then
    Txt_CFOP_prod = IIf(IsNull(TBOSC!ID_CFOP), "", TBOSC!ID_CFOP)
    Txt_natureza_operacao_prod = IIf(IsNull(TBOSC!Txt_descricao), "", TBOSC!Txt_descricao)
    If TBOSC!retorno = True Then chkRetorno.Value = 1 Else chkRetorno.Value = 0
End If
TBOSC.Close

Txt_data_retorno = "__/__/____"
N_item = ""
txtdesctecnica.Text = ""
txtQuantidade.Text = ""
txtDesconto.Text = 0
txtvalordesconto.Text = ""
txtvalorunitariodesc.Text = ""
txtEspecificacoes.Text = ""
Txt_observacoes_prod = ""
Cmb_prioridade = "Normal"
Txt_observacoes_fat_prod = ""
If Novo_PI1 = True And Chk_obs_faturamento_prod.Value = 1 Then
    Set TBOSC = CreateObject("adodb.recordset")
    TBOSC.Open "Select Obs_faturamento from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'P' and Obs_faturamento IS NOT NULL order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBOSC.EOF = False Then
        Txt_observacoes_fat_prod = TBOSC!Obs_faturamento
    End If
    TBOSC.Close
End If

Chk_antecipacao.Value = 0
Chk_faturamento_parcial.Value = 0
chkNovo_projeto.Value = 0
Chk_utiliza_mat_consignado.Value = 0
txtespessura = ""
txtLargura = ""
txtComprimento = ""
txtDureza = ""
txtvalor_total.Text = ""
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
cmbfamilia.ListIndex = -1
txtdbl_valoripi.Text = ""
txtint_icms.Text = ""
txtInt_ipi.Text = ""
txtvalor_icms.Text = ""
txtinspecao = ""
txtembalagem = ""
txtGravacao = ""
txtComissao = ""
CodigoLista = 0
If Produtos2 = True Then ProcCarregaComboProduto
'ProcBloqueiaLibera_Validacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboProduto()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Vendas = 'True'", False
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboMoeda()
On Error GoTo tratar_erro

With cmbMoeda
    .Clear
    Set TBFamilia = CreateObject("adodb.recordset")
    TBFamilia.Open "Select * from moeda", Conexao, adOpenKeyset, adLockOptimistic
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

Sub ProcLimparServicos(Servicos2 As Boolean)
On Error GoTo tratar_erro

If Servicos2 = False Then
    IDAnalise_servico = 0
    Txt_analise1 = ""
    txtid_servico = 0
    txtcodservico.Text = ""
    txtpcclienteserv = ""
    Caminho_PC_serv_PI = ""
    txtPrazo_Servico = ""
    mskprazoservico.Text = "__/__/____"
    If Novo_PI2 = True Then
        If Chk_PC_serv.Value = 1 Then
            Set TBOSC = CreateObject("adodb.recordset")
            TBOSC.Open "Select PCcliente, Caminho_PCCliente from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'S' and PCcliente IS NOT NULL order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBOSC.EOF = False Then
                 txtpcclienteserv = TBOSC!PCCliente
                 Caminho_PC_serv_PI = IIf(IsNull(TBOSC!Caminho_PCCliente), "", TBOSC!Caminho_PCCliente)
            End If
            TBOSC.Close
        End If
        If Chk_prazo_serv.Value = 1 Then
            Set TBOSC = CreateObject("adodb.recordset")
            TBOSC.Open "Select Prazofinaldias, Prazofinal from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'S' and (Prazofinaldias IS NOT NULL or prazofinal IS NOT NULL) order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBOSC.EOF = False Then
                 If Vendas_PI = True Then mskprazoservico = TBOSC!PrazoFinal Else txtPrazo_Servico = TBOSC!prazofinaldias
            End If
            TBOSC.Close
        End If
    End If
End If
cmbreferencia_serv.Clear
txtReferencia_serv = ""
Txt_ID_CFOP_serv = ""
Txt_CFOP_serv = ""
Txt_natureza_operacao_serv = ""
Permitido = False
If Novo_PI2 = True Then
    If Chk_CFOP_serv.Value = 1 Then
        Set TBOSC = CreateObject("adodb.recordset")
        TBOSC.Open "Select ID_CFOP from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'S' and ID_CFOP IS NOT NULL order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBOSC.EOF = False Then
            Txt_ID_CFOP_serv = TBOSC!ID_CFOP
            Permitido = True
        End If
        TBOSC.Close
    End If
    If Permitido = False Then
        'Verifica CFOP vinculada ao cliente
        Set TBOSC = CreateObject("adodb.recordset")
        TBOSC.Open "Select IDCFOP FROM Clientes_DadosComerciais where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBOSC.EOF = False Then
            Txt_ID_CFOP_serv = IIf(IsNull(TBOSC!IDCFOP), "", TBOSC!IDCFOP)
        End If
    End If
End If
Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select ID_CFOP, Txt_descricao FROM tbl_NaturezaOperacao where IDCountCfop = " & IIf(Txt_ID_CFOP_serv = "", 0, Txt_ID_CFOP_serv), Conexao, adOpenKeyset, adLockOptimistic
If TBOSC.EOF = False Then
    Txt_CFOP_serv = IIf(IsNull(TBOSC!ID_CFOP), "", TBOSC!ID_CFOP)
    Txt_natureza_operacao_serv = IIf(IsNull(TBOSC!Txt_descricao), "", TBOSC!Txt_descricao)
End If
TBOSC.Close

optnovoservico.Value = 0
OPTnovoservicoman.Value = 0
txtqtservico.Text = ""
txtRev_serv.Text = ""
Cmb_cidade_servico.ListIndex = -1
txtunservico.ListIndex = -1
Cmb_un_com_serv.ListIndex = -1
txtdescservico.Text = ""
Chk_servico_executado_cliente.Value = 0
txtdesccomservico.Text = ""
Chk_antecipacao_serv.Value = 0
Chk_faturamento_parcial_serv.Value = 0
Chk_utiliza_mat_consignado_serv.Value = 0
cmbfamiliaservico.ListIndex = -1
txtvlrunitservico.Text = ""
txtdesconto2.Text = 0
txtvalordesconto2.Text = ""
txtvalorunitariodesc2.Text = ""
ProcCarregaISSQN
txtvlrISS.Text = ""
txtvlrtotalservico.Text = ""
txtObs_serv = ""

Txt_observacoes_fat_serv = ""
If Novo_PI2 = True And Chk_obs_faturamento_serv.Value = 1 Then
    Set TBOSC = CreateObject("adodb.recordset")
    TBOSC.Open "Select Obs_faturamento from vendas_carteira where cotacao = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'S' and Obs_faturamento IS NOT NULL order by Codigo desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBOSC.EOF = False Then
        Txt_observacoes_fat_serv = TBOSC!Obs_faturamento
    End If
    TBOSC.Close
End If

txtComissaoServ = ""
CodigoLista1 = 0
If Servicos2 = True Then ProcCarregaComboServico
'ProcBloqueiaLibera_ValidacaoServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboServico()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamiliaservico, "familia <> 'Null' and Vendas = 'True'", False
ProcCarregaComboUnidade txtunservico, False
ProcCarregaComboUnidade Cmb_un_com_serv, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparTotais()
On Error GoTo tratar_erro

txt_BaseICMS = "0,00"
txt_vlrICMS = "0,00"
txt_baseICMSs = "0,00"
txt_ICMSs = "0,00"
txt_vlrtotalprod = "0,00"
txttotalservicos = "0,00"
txtTotaldesconto = "0,00"
txt_TotalIPI = "0,00"
txt_ValorNota = "0,00"
txttotalproposta = "0,00"
TxtTotalFrete.Text = "0,00"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
 
txtId.Text = 0
txtCotacao.Text = ""
txtrevisao.Text = ""
txtDtValidacao = ""
txtRespValidacao = ""
txt_datamodificado.Text = ""
txtStatus.Text = ""
txtIDcliente.Text = ""
txtCliente.Text = ""
txttipocliente.ListIndex = -1
txtRemetente.Text = ""
txtdepartamento.Text = ""
txttelefone.Text = ""
txtFax.Text = ""
txtEmail.Text = ""
cmbTipo_endereco.ListIndex = -1
txtendereco = ""
txtNumero = ""
txtComplemento = ""
cmbTipo_bairro.ListIndex = -1
txtBairro = ""
txtCidade.Text = ""
txtuf.ListIndex = -1
txtRef.Text = ""
txtreferente.Text = ""
txtDatavendas = ""
txtDatavendas_PI = Format(Date, "dd/mm/yyyy")
txtVI.Text = ""
txtvend_Int.Text = ""
txtVE.Text = ""
txtVend_Ext.Text = ""
txtregiao.Text = ""
txt_dataelaborado.Value = Date
txtResponsavel.Text = pubUsuario
txt_observacoes.Text = ""
CodigoLista2 = 0
TabelaSN_PI = 0
RegimeEmpresa_PI = 0
Caption = "Administrativo - Vendas - " & IIf(Vendas_Proposta = True, "Proposta comercial", "Pedido interno")
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtCotacao.Text = "" Or Novo_PI = True Then
    SSTab1.Tab = 0
        Frame1(0).Visible = False
        Frame1(2).Visible = False
    
    Exit Sub
End If

ProcCorrigeForm
'PBLista.Visible = True

Select Case SSTab1.Tab
    Case 0: 'Pedido interno dados gerais
        If Lista.Visible = True Then Lista.SetFocus
        Frame1(0).Visible = False
        Frame1(2).Visible = False
        LiberarAlteracao = False
    Case 1: 'Pedido interno dados comerciais
        'PBLista.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        cmdBotao_DadosComerciais(0).SetFocus
        ProcPuxaDadosComercial
        Frame1(0).Visible = False
        Frame1(2).Visible = False
       
    Case 2: 'Lista de produtos
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Listprod.SetFocus
        ProcAtualizalistaProdutos (1)
        Frame1(0).Visible = True
        Frame1(2).Visible = True
        Prod = True
        LiberarAlteracao = False
    Case 3: 'Lista de serviços
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ListaServicos.SetFocus
        ProcAtualizalistaServicos (1)
        Frame1(0).Visible = True
        Frame1(2).Visible = True
        Prod = False
        
    Case 4: 'Escopo de fornecimento
        'PBLista.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        txtEscopo.SetFocus
        ProcCarregaEscopoForn
        Frame1(0).Visible = False
        Frame1(2).Visible = False
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeForm()
On Error GoTo tratar_erro

If Vendas_PI = True Then
    txtStatus.Width = txtIDcliente.Left - txtStatus.Left
 '   Label1(56).Left = txtStatus.Left + (txtStatus.Width / 2) - (Label1(56).Width / 2)
End If
Select Case SSTab1.Tab
''    Case 0: Frame1(0).Top = Frame1(1).Top + Frame1(1).Height
'    Case 1: Frame1(0).Top = Frame1(4).Top + Frame1(4).Height
'    Case 2: Frame1(0).Top = Frame1(6).Top + Frame1(6).Height
'    Case 3: Frame1(0).Top = Frame1(8).Top + Frame1(6).Height
'    Case 4: Frame1(0).Top = Frame1(9).Top + Frame1(9).Height
End Select
'Frame1(2).Top = Frame1(0).Top

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_PI = True Then
    USMsgBox ("Salve a proposta antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    Permitido = False
    Exit Sub
End If
If txtIDcliente = "" Or txtCliente = "" Then
    USMsgBox ("Informe o cliente antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    If txtIDcliente = "" Then txtIDcliente.SetFocus Else txtCliente.SetFocus
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEscopoForn()
On Error GoTo tratar_erro

txtEscopo = ""
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM vendas_comercial WHERE cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    txtEscopo = IIf(IsNull(TBCotacao!Escopo_fornecimento), "", TBCotacao!Escopo_fornecimento)
End If
TBCotacao.Close

txtObsAnalise = ""
ObsAnalise = ""
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select VAS.* from vendas_carteira VC INNER JOIN Vendas_analise_setores VAS ON VC.idanalise = VAS.IDanalise where VC.cotacao = " & txtId.Text & " and (VAS.Setor = 'ENGENHARIA' or VAS.Setor = 'QUALIDADE') order by VAS.IDanalise, VAS.Codinterno", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        If ObsAnalise = "" Then
            ObsAnalise = TBAbrir!Texto & " - " & TBAbrir!Un & " - " & Format(TBAbrir!Qtde, "###,##0.0000") & " - " & Format(TBAbrir!VlrUnit, "###,##0.0000") & " - " & Format(TBAbrir!vlrTotal, "###,##0.00")
        Else
            ObsAnalise = ObsAnalise & vbCrLf & TBAbrir!Texto & " - " & TBAbrir!Un & " - " & Format(TBAbrir!Qtde, "###,##0.0000") & " - " & Format(TBAbrir!VlrUnit, "###,##0.0000") & " - " & Format(TBAbrir!vlrTotal, "###,##0.00")
        End If
        TBAbrir.MoveNext
    Loop
    txtObsAnalise = ObsAnalise
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcgravarItem()
On Error GoTo tratar_erro

If txtStatus <> "VENDIDA PARCIAL" And txtStatus <> "REVISADA" And txtStatus <> "FATURADA" And txtStatus <> "FATURADA PARCIAL" Then
    DataVendasTexto = ""
    Select Case txtStatus
        Case "ABERTA EM ANALISE": StatusTexto = "ABERTA EM ANALISE"
        Case "VENDIDA":
            StatusTexto = "VENDIDA"
            DataVendasTexto = txtDatavendas_PI
        Case "CANCELADA": StatusTexto = "CANCELADO"
        Case "PERDIDA P/ PRAZO": StatusTexto = "PERDIDO P/ PRAZO"
        Case "PERDIDA P/ PREÇO": StatusTexto = "PERDIDO P/ PREÇO"
        Case "PORTAL ELETRONICO": StatusTexto = "PORTAL ELETRONICO"
    End Select
    Conexao.Execute "UPDATE vendas_carteira Set Liberacao = '" & StatusTexto & "', Datavendas = '" & IIf(DataVendasTexto = "", Null, DataVendasTexto) & "' where cotacao = " & txtId
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaDesconto()
On Error GoTo tratar_erro

If txtvalorunitario.Text <> "" And txtQuantidade <> "" Then
    If IsNumeric(txtvalorunitario.Text) = True Then
        a = Format(txtvalorunitario.Text, "###,##0.0000000000")
        c = IIf(txtDesconto = "", 0, txtDesconto)
        D = (a * c) / 100
        E = txtQuantidade.Text
'=============================================================================
        txtvalordesconto.Text = Format((D * E), "###,##0.0000000000")
'=============================================================================
        txtvalorunitariodesc.Text = Format(a - D, "###,##0.0000000000")
        ProcAtualizavalores
    End If
Else
    txtvalordesconto = "0,00000"
    txtvalorunitariodesc = IIf(txtvalorunitario = "", "0,00000", txtvalorunitario)
    txtvalor_total = "0,00"
    txtdbl_valoripi = "0,00"
    txtvalor_icms = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaDesconto2()
On Error GoTo tratar_erro

If txtvlrunitservico.Text <> "" And txtqtservico <> "" And txtdesconto2 <> "" Then
    If IsNumeric(txtvlrunitservico.Text) = True Then
        a = Format(txtvlrunitservico.Text, "###,##0.0000000000")
        c = IIf(txtdesconto2 = "", 0, txtdesconto2)
        D = (a * c) / 100
        txtvalordesconto2.Text = Format(D, "###,##0.0000000000")
        txtvalorunitariodesc2.Text = Format(a - D, "###,##0.0000000000")
    End If
Else
    txtvalordesconto2 = "0,00000"
    txtvalorunitariodesc2 = IIf(txtvlrunitservico = "", "0,00000", txtvlrunitservico)
    txtvlrtotalservico = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaTotais()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
ProcLimparTotais
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_proposta where cotacao = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    txt_BaseICMS.Text = IIf(IsNull(TBVendas!dbl_Base_ICMS), "0,00", Format(TBVendas!dbl_Base_ICMS, "###,##0.00"))
    txt_vlrICMS.Text = IIf(IsNull(TBVendas!dbl_Valor_ICMS), "0,00", Format(TBVendas!dbl_Valor_ICMS, "###,##0.00"))
    txt_baseICMSs.Text = IIf(IsNull(TBVendas!dbl_Base_ICMS_Subst), "0,00", Format(TBVendas!dbl_Base_ICMS_Subst, "###,##0.00"))
    txt_ICMSs.Text = IIf(IsNull(TBVendas!dbl_Valor_ICMS_Subst), "0,00", Format(TBVendas!dbl_Valor_ICMS_Subst, "###,##0.00"))
    txt_vlrtotalprod.Text = IIf(IsNull(TBVendas!dbl_Valor_Total_Produtos), "0,00", Format(TBVendas!dbl_Valor_Total_Produtos, "###,##0.00"))
    txttotalservicos.Text = IIf(IsNull(TBVendas!dbl_valor_total_servicos), "0,00", Format(TBVendas!dbl_valor_total_servicos, "###,##0.00"))
    txtTotaldesconto.Text = IIf(IsNull(TBVendas!TotalDesconto), "0,00", Format(TBVendas!TotalDesconto, "###,##0.00"))
    txt_ValorNota.Text = IIf(IsNull(TBVendas!SubTotal), "0,00", Format(TBVendas!SubTotal, "###,##0.00"))
    txt_TotalIPI.Text = IIf(IsNull(TBVendas!dbl_Valor_Total_IPI), "0,00", Format(TBVendas!dbl_Valor_Total_IPI, "###,##0.00"))
    txttotalproposta.Text = IIf(IsNull(TBVendas!dbl_valor_total), "0,00", Format(TBVendas!dbl_valor_total, "###,##0.00"))
    TxtTotalFrete.Text = IIf(IsNull(TBVendas!VTotalfrete), "0,00", Format(TBVendas!VTotalfrete, "###,##0.00"))
End If
TBVendas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCarregaISSQN()
On Error GoTo tratar_erro

If Chk_servico_executado_cliente.Value = 0 Then
    ProcVerifImpostosEmpresa Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False, "", False, 0, True, TabelaSN_PI, 0
    txtiss = ISS_Serv
    With Cmb_cidade_servico
        .ListIndex = -1
        .Locked = False
        .TabStop = True
    End With
Else
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ISSQN from Clientes where IDCliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        txtiss = IIf(IsNull(TBFIltro!ISSQN), 0, TBFIltro!ISSQN)
    End If
    TBFIltro.Close
    If cmbCidade <> "" Then
        With Cmb_cidade_servico
            .Text = cmbCidade
            .Locked = True
            .TabStop = False
        End With
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame1(10).Enabled = False
Frame1(12).Enabled = False
Frame1(11).Enabled = False
'ProcLimparComercial
ProcLimparProdutos False
ProcLimparServicos False
ProcLimparTotais
Novo_PI1 = False
Novo_PI2 = False
Novo_PI3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaClientes()
On Error GoTo tratar_erro

If Vendas_Proposta = True Then
    txtCliente.Locked = False
    cmbTipo_endereco.Locked = False
    txtendereco.Locked = False
    txtNumero.Locked = False
    txtComplemento.Locked = False
    cmbTipo_bairro.Locked = False
    txtBairro.Locked = False
    txtuf.Locked = False
    cmbCidade.Locked = False
End If
INNERJOINTEXTO = "clientes C"
TextoFiltro = ""
If txtVend_Ext <> "" Then
    INNERJOINTEXTO = "(clientes C LEFT JOIN vendas_vendedores VV ON VV.N_Vendedor = " & txtVE & ") LEFT JOIN Vendas_Vendedores_Clientes VVC ON VVC.IDVendedor = VV.ID and VVC.IDCliente = C.IDCliente"
    TextoFiltro = " and (VV.Bloquear_venda_cliente = 'True' and VVC.IDCliente IS NOT NULL or VV.Bloquear_venda_cliente = 'False')"
End If
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select C.* from " & INNERJOINTEXTO & " where C.IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and C.status <> 'Bloqueado'" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    txtIDcliente.Text = TBFI!IDCliente
    txtCliente.Text = IIf(IsNull(TBFI!NomeRazao), "", TBFI!NomeRazao)
    txttipocliente = IIf(IsNull(TBFI!Tipo), "", (TBFI!Tipo))
    txtEmail = IIf(IsNull(TBFI!Email), "", TBFI!Email)
    txttelefone = IIf(IsNull(TBFI!Tel01), "", TBFI!Tel01)
    txtFax = IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
    NomeCampo = "o tipo do endereço"
    
    If IsNull(TBFI!Tipo_endereco) = False And TBFI!Tipo_endereco <> "" Then
    cmbTipo_endereco.Text = TBFI!Tipo_endereco
    End If
    
    txtendereco.Text = IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco)
    txtNumero = IIf(IsNull(TBFI!Numero), "", TBFI!Numero)
    txtComplemento.Text = IIf(IsNull(TBFI!complemento), "", TBFI!complemento)
    NomeCampo = "o tipo do bairro"
    
    If IsNull(TBFI!Tipo_bairro) = False And TBFI!Tipo_bairro <> "" Then
    cmbTipo_bairro.Text = TBFI!Tipo_bairro
    End If
    
    txtBairro.Text = IIf(IsNull(TBFI!Bairro), "", TBFI!Bairro)
    NomeCampo = "o estado"
    If IsNull(TBFI!UF) = False And TBFI!UF <> "" And TBFI!UF <> txtuf Then txtuf.Text = TBFI!UF
    If txtuf = "EX" Then
        txtCidade.Visible = True
        cmbCidade.Visible = False
        txtCidade.Text = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
    Else
        txtCidade.Visible = False
        cmbCidade.Visible = True
        If IsNull(TBFI!Cidade) = False And TBFI!Cidade <> "" Then cmbCidade.Text = TBFI!Cidade
    End If
1:
    txt_observacoes = IIf(IsNull(TBFI!txt_observacoes), "", TBFI!txt_observacoes)
    
    txtCliente.Locked = True
    cmbTipo_endereco.Locked = True
    txtendereco.Locked = True
    txtNumero.Locked = True
    txtComplemento.Locked = True
    cmbTipo_bairro.Locked = True
    txtBairro.Locked = True
    txtuf.Locked = True
    cmbCidade.Locked = True
End If
TBFI.Close
                
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste cliente."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCliente()
On Error GoTo tratar_erro

txtCliente.Text = ""
txttipocliente.ListIndex = -1
txtRemetente.Text = ""
txtdepartamento.Text = ""
txttelefone.Text = ""
txtFax.Text = ""
txtEmail.Text = ""
cmbTipo_endereco.ListIndex = -1
txtendereco = ""
txtNumero = ""
txtComplemento = ""
cmbTipo_bairro.ListIndex = -1
txtBairro = ""
txtCidade.Text = ""
txtuf.ListIndex = -1
                
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorDesconto()
On Error GoTo tratar_erro

Dim Valorunitario As Double
Dim ValorDesconto As Double
Dim Quantida As Double
Dim TotalDesconto As Double
Dim PDesconto As Double
Dim ValorAntigo As Double


If txtvalorunitario.Text <> "" And txtQuantidade <> "" Then
    If IsNumeric(txtvalorunitario.Text) = True Then
        Valorunitario = txtvalorunitario.Text
        ValorAntigo = txtvalorunitario.Text
        Quantida = txtQuantidade.Text
        If txtvalordesconto.Text <> "" Then
        ValorDesconto = txtvalordesconto / Quantida
        End If
        Valorunitario = Valorunitario - ValorDesconto
        If ValorDesconto <> 0 And ValorAntigo <> 0 Then
        PDesconto = ((ValorDesconto) / ValorAntigo) * 100
        End If
       
        
        txtDesconto.Text = Format(PDesconto, "###,##0.0000000000")
        txtvalorunitariodesc.Text = Format(Valorunitario, "###,##0.0000000000")
    Else
        Exit Sub
    End If
    ProcAtualizavalores
Else
    txtvalordesconto = "0,00000"
    txtvalorunitariodesc = IIf(txtvalorunitario = "", "0,00000", txtvalorunitario)
    txtvalor_total = "0,00"
    txtdbl_valoripi = "0,00"
    txtvalor_icms = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorDescontoServ()
On Error GoTo tratar_erro

If txtvlrunitservico.Text <> "" And txtqtservico <> "" Then
    If IsNumeric(txtvlrunitservico.Text) = True Then
        quantestoque = txtvlrunitservico.Text
        QuantSolicitado = IIf(txtvalordesconto2 = "", 0, txtvalordesconto2)
        If quantestoque <> 0 Then QuantEmpenho = (QuantSolicitado * 100) / quantestoque Else QuantEmpenho = 0
        txtdesconto2.Text = QuantEmpenho
        txtvalorunitariodesc2.Text = Format(quantestoque - QuantSolicitado, "###,##0.0000000000")
    Else
        Exit Sub
    End If
    ProcCalculaValoresServicos
Else
    txtvalordesconto2 = "0,00000"
    txtvalorunitariodesc2 = IIf(txtvlrunitservico = "", "0,00000", txtvlrunitservico)
    txtvlrtotalservico = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro
'Debug.print ButtonIndex

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcCopiar
    Case 8: ProcAbrirCheckList
    Case 9: ProcRevisao
    Case 10: ProcEmitirPI
    Case 11: ProcCancelaPI
    Case 12: ProcImpostos
    Case 13: ProcStatus
    Case 14: ProcValidarRegistros Lista, IIf(Vendas_Proposta = True, "Vendas/Proposta comercial", "Vendas/Pedido interno")
    Case 15: ProcImportarExcel
    Case 16: procAtualiza
    Case 18: frmSenha.Show 1 'ProcAjuda
    Case 19: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirCheckList()
On Error GoTo tratar_erro

    If txtId.Text <> "" And txtId.Text <> 0 Then
        frmVendas_PI_CheckList_Compras.Show
        StrSql = ""
    Else
        USMsgBox "Escolha um pedido de venda para verificar as necessidade de compras do(s) item(ns) vendido(s)", vbInformation, "CAPRIND v5.0"
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
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

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_prod
    Case 2: ProcSalvar_prod
    Case 3: ProcExcluir_prod
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcEstrutura
    Case 8: ProcEmitirPI_prod
    Case 9: ProcCancelaPI_prod
    Case 10: ProcComposicao_prod
    Case 11: ProcStatus
    Case 12: ProcAlteracoes
    Case 13:
    
    If Id_Item <> 0 Then
    frmVendas_PI_Empenhos.Show 1
    End If
    
    Case 15: frmSenha.Show 1
    Case 16: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovo_serv
    Case 2: ProcSalvar_serv
    Case 3: ProcExcluir_Serv
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcEmitirPI_serv
    Case 8: ProcCancelaPI_serv
    Case 9: ProcComposicao_serv
    Case 10: ProcStatus
    Case 11: ProcAlteracoes
    Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
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

Sub ProcFinanceiro()
On Error GoTo tratar_erro

If txtStatus <> "VENDIDA" And txtStatus <> "VENDIDA PARCIAL" Then
    USMsgBox ("Só é permitido enviar para o financeiro " & IIf(Vendas_PI = True, "pedido interno", "proposta comercial") & " com o status vendida ou vendida parcial."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "enviar para o financeiro"
If txttotalproposta = "" Or txttotalproposta = "0,00" Then
    NomeCampo = "o valor total " & IIf(Vendas_PI = True, "do pedido interno", "da proposta comecial")
    ProcVerificaAcao
    Exit Sub
End If
frmVendas_propostaII_MenuFinanceiro.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro
'Debug.print StrSql_PI_Localizar

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSql_PI_Localizar = "" Then Exit Sub
Set TBLISTA_Vendas_PI = CreateObject("adodb.recordset")
TBLISTA_Vendas_PI.Open StrSql_PI_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_PI.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Cotacao = 0
Lista.ListItems.Clear
TBLISTA_Vendas_PI.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Vendas_PI.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_PI.PageSize
ContadorReg = 1
'PBLista.Min = 0
'PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_PI.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_PI.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_PI.PageSize)
'PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendas_PI.EOF = False And (ContadorReg <= TamanhoPagina)
StrSql = "select vca.cotacao, sum(vca.Preco_lote) as Total, vco.Valor_frete as Valor_Frete from vendas_carteira as vca inner join vendas_comercial as vco  ON vca.cotacao = vco.cotacao where vca.cotacao = " & TBLISTA_Vendas_PI!Cotacao & " and (Liberacao = 'VENDIDA' OR Liberacao = 'FATURADO' OR Liberacao = 'CANCELADO' OR Liberacao = 'FATURADO PARCIAL') group by vca.cotacao, vco.Valor_frete"
'Debug.print StrSql

    With Lista.ListItems
        .Add , , TBLISTA_Vendas_PI!Cotacao
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_PI!Data), "", Format(TBLISTA_Vendas_PI!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = TBLISTA_Vendas_PI!Ncotacao
        .Item(.Count).SubItems(3) = TBLISTA_Vendas_PI!Revisao
        
        If UseNomeFantasia = True Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select nomefantasia from clientes where idCliente = " & TBLISTA_Vendas_PI!IDCliente, Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Lista.ColumnHeaders.Item(5).Text = "Nome Fantasia"
                .Item(.Count).SubItems(4) = IIf(IsNull(TBClientes!NomeFantasia), "", Trim(TBClientes!NomeFantasia))
            End If
            TBClientes.Close
        Else
            Lista.ColumnHeaders.Item(5).Text = "Razão social"
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_PI!Cliente), "", Trim(TBLISTA_Vendas_PI!Cliente))
        End If
        
        Set TBTotaisnota = CreateObject("adodb.recordset")
        TBTotaisnota.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
        If TBTotaisnota.EOF = False Then
        'Total
        .Item(.Count).SubItems(5) = Format((IIf(IsNull(TBTotaisnota!Total), "0", TBTotaisnota!Total) + IIf(IsNull(TBLISTA_Vendas_PI!VTotalfrete), "0", TBLISTA_Vendas_PI!VTotalfrete)), "###,##0.00")
        End If
        TBTotaisnota.Close
        
        'ProcGravarTotais TBLISTA_Vendas_PI!Cotacao
        '.Item(.Count).SubItems(5) = Format(SubTotal + SumIPI + TotalICMSCST, "###,##0.00")
        
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Vendas_PI!status), "", TBLISTA_Vendas_PI!status)
        If Vendas_Proposta = True Then
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Vendas_PI!DtValidacao) = False, "Sim", "Não")
        Else
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Vendas_PI!DtValidacaoPI) = False, "Sim", "Não")
        End If
    End With
    TBLISTA_Vendas_PI.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    'PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Vendas_PI.RecordCount
If TBLISTA_Vendas_PI.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Vendas_PI.PageCount
ElseIf TBLISTA_Vendas_PI.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Vendas_PI.PageCount & " de: " & TBLISTA_Vendas_PI.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Vendas_PI.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_PI.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcOrganizaFormPI_Proposta()
On Error GoTo tratar_erro

If Vendas_PI = True Then
    Formulario = "Vendas/Pedido interno"
    Caption = "Administrativo - Vendas - Pedido interno"
    txtDatavendas_PI = Date
    Label1(15).Caption = "Pedido"
    Label1(27).Caption = "Prazo"
    Label1(33).Caption = "Prazo"
    Label1(80).Caption = "Total pedido"
    cmdstatus.Visible = False
    With txtCliente
        .Locked = True
        .TabStop = False
    End With
    With cmbTipo_endereco
        .Locked = True
        .TabStop = False
    End With
    With txtendereco
        .Locked = True
        .TabStop = False
    End With
    With txtNumero
        .Locked = True
        .TabStop = False
    End With
    With txtComplemento
        .Locked = True
        .TabStop = False
    End With
    With cmbTipo_bairro
        .Locked = True
        .TabStop = False
    End With
    With txtBairro
        .Locked = True
        .TabStop = False
    End With
    With txtuf
        .Locked = True
        .TabStop = False
    End With
    With cmbCidade
        .Locked = True
        .TabStop = False
    End With
    txtDatavendas.Visible = False
    txtDatavendas_PI.Visible = True
    txtPrazo_Produto.Visible = False
    mskprazo.Visible = True
    imgCalendario.Visible = True
    Chk_PC_prod.Visible = True
    With txtpccliente
        .Locked = False
        .TabStop = True
    End With
    Chk_PC_serv.Visible = True
    With txtpcclienteserv
        .Locked = False
        .TabStop = True
    End With
    txtPrazo_Servico.Visible = False
    mskprazoservico.Visible = True
    ImgCalendario1.Visible = True
    USToolBar1.ButtonState(10) = 5
    USToolBar1.ButtonState(11) = 5
    USToolBar3.ButtonState(8) = 5
    USToolBar4.ButtonState(7) = 5
    Lista.ColumnHeaders(3).Text = "Pedido"
Else
    Formulario = "Vendas/Proposta comercial"
    Caption = "Administrativo - Vendas - Proposta comercial"
    USToolBar1.ButtonState(14) = 5
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcComposicao_prod()
On Error GoTo tratar_erro

If txtid_produto = 0 Then Exit Sub
PI_Produtos = True
PI_Servicos = False
frmVendas_PI_composicao_prodserv.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcComposicao_serv()
On Error GoTo tratar_erro

If txtid_servico = 0 Then Exit Sub
PI_Produtos = False
PI_Servicos = True
frmVendas_PI_composicao_prodserv.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaLibera_Validacao()
On Error GoTo tratar_erro

If txtDtValidacao = "" Then Permitido = True Else Permitido = False
ProcLibBlocTxt txtNomenclatura, Permitido
ProcLibBlocTxt txtRev_cod, Permitido
ProcLibBlocCmb cmbReferencia, Permitido
ProcLibBlocTxt txtdesctecnica, Permitido
ProcLibBlocTxt txtEspecificacoes, Permitido
ProcLibBlocTxt txtespessura, Permitido
ProcLibBlocTxt txtLargura, Permitido
ProcLibBlocTxt txtComprimento, Permitido
'ProcLibBlocTxt txtDureza, Permitido
ProcLibBlocCmb Cmb_CST_ICMS, Permitido
ProcLibBlocCmb cmbun, Permitido
ProcLibBlocCmb Cmb_un_com, Permitido
ProcLibBlocCmb cmbfamilia, Permitido
ProcLibBlocTxt txtQuantidade, Permitido
ProcLibBlocTxt txtvalorunitario, Permitido
ProcLibBlocTxt txtComissao, Permitido
If Permitido = True Then
    cmdfiltrar.Enabled = True
    cmdlistaproduto.Enabled = True
    Cmd_localizar_CFOP_prod.Enabled = True
    Cmd_limpar_CFOP.Enabled = True
    Frame1(5).Enabled = True
    Chk_desc.Enabled = True
    Chk_valor_desc.Enabled = True
    cmdCF.Enabled = True
    Cmd_limpar_CF.Enabled = True
    Cmd_analise.Enabled = True
    chkRetorno.Enabled = True
Else
    cmdfiltrar.Enabled = False
    cmdlistaproduto.Enabled = False
    Cmd_localizar_CFOP_prod.Enabled = False
    Cmd_limpar_CFOP.Enabled = False
    Frame1(5).Enabled = False
    Chk_desc.Enabled = False
    Chk_valor_desc.Enabled = False
    cmdCF.Enabled = False
    Cmd_limpar_CF.Enabled = False
    Cmd_analise.Enabled = False
    chkRetorno.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaLibera_ValidacaoServ()
On Error GoTo tratar_erro

If txtDtValidacao = "" Then Permitido = True Else Permitido = False
ProcLibBlocTxt txtcodservico, Permitido
ProcLibBlocTxt txtRev_serv, Permitido
ProcLibBlocCmb cmbreferencia_serv, Permitido
ProcLibBlocTxt txtdescservico, Permitido
ProcLibBlocTxt txtdesccomservico, Permitido
ProcLibBlocCmb txtunservico, Permitido
ProcLibBlocCmb Cmb_un_com_serv, Permitido
ProcLibBlocCmb cmbfamiliaservico, Permitido
ProcLibBlocTxt Txt_CFOP_serv, Permitido
ProcLibBlocTxt Txt_natureza_operacao_serv, Permitido
ProcLibBlocTxt txtqtservico, Permitido
ProcLibBlocTxt txtvlrunitservico, Permitido
ProcLibBlocTxt txtComissaoServ, Permitido
ProcLibBlocCmb Cmb_cidade_servico, Permitido
ProcLibBlocTxt txtiss, Permitido
If Permitido = True Then
    cmdfiltrar_serv.Enabled = True
    cmdlistaservicos.Enabled = True
    Cmd_localizar_CFOP_serv.Enabled = True
    Cmd_limpar_CFOP_serv.Enabled = True
    Frame1(7).Enabled = True
    Cmd_analise1.Enabled = True
    Chk_desc2.Enabled = True
    Chk_valor_desc2.Enabled = True
Else
    cmdfiltrar_serv.Enabled = False
    cmdlistaservicos.Enabled = False
    Cmd_localizar_CFOP_serv.Enabled = False
    Cmd_limpar_CFOP_serv.Enabled = False
    Frame1(7).Enabled = False
    Cmd_analise1.Enabled = False
    Chk_desc2.Enabled = False
    Chk_valor_desc2.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Compras_Pedido = False
Plano_centro_de_custo = False

Permitido = False
If SSTab1.Tab = 0 Then
    Sit_REG = 1
    TextoPadrao = IIf(Vendas_PI = True, "o(s) pedido(s)", "a(s) proposta(s)")
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then Permitido = True
        Next InitFor
    End With
ElseIf SSTab1.Tab = 2 Then
        Sit_REG = 2
        TextoPadrao = "o(s) produtos(s)"
        With Listprod
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then Permitido = True
            Next InitFor
        End With
    Else
        Sit_REG = 3
        TextoPadrao = "o(s) serviço(s)"
        With ListaServicos
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then Permitido = True
            Next InitFor
        End With
End If
If Permitido = False Then
    USMsgBox ("Informe " & TextoPadrao & " antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

frmCompras_pedido_cancelar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlteracoes()
On Error GoTo tratar_erro

Permitido = True
If SSTab1.Tab = 2 Then
    TextoPadrao = "produtos"
    Sit_REG = 1
    If txtid_produto = 0 Then Permitido = False
Else
    TextoPadrao = "serviço"
    Sit_REG = 2
    If txtid_servico = 0 Then Permitido = False
End If
If Permitido = False Then
    USMsgBox ("Informe o " & TextoPadrao & " antes de cadastrar as alterações."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Compras_Pedido = False
frmVendas_PI_alteracoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImportarExcel()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmVendas_PI_importar_excel.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifAltCodQtde(Prod As Boolean, AltCod As Boolean) As Boolean
On Error GoTo tratar_erro

FunVerifAltCodQtde = True
TextoMsg = IIf(AltCod = True, "o código interno", "a quantidade")
TextoMsg1 = IIf(Prod = True, "produto", "serviço")

If TBCotacao!Liberacao = "FATURADO" Then
    USMsgBox "Não é permitido alterar , pois o status do item é " & TBCotacao!Liberacao & ", vbExclamation, CAPRIND v5.0"
    TBCotacao.Close
    FunVerifAltCodQtde = False
    Exit Function
End If

'Set TBAbrir = CreateObject("adodb.recordset")
'TBAbrir.Open "Select Idmateriaprima from Producaomaterial where ID_carteira = " & TBCotacao!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
'If TBAbrir.EOF = False Then
'    usMsgbox ("Não é permitido alterar " & TextoMsg & " deste " & TextoMsg1 & ", pois o mesmo já gerou necessidade de compras."), vbExclamation, "CAPRIND v5.0"
'    TBCotacao.Close
'    TBAbrir.Close
'    FunVerifAltCodQtde = False
'    Exit Function
'End If
'TBAbrir.Close
'Set TBAbrir = CreateObject("adodb.recordset")
'TBAbrir.Open "Select ID from Producao_pedidos where IDcarteira = " & TBCotacao!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
'If TBAbrir.EOF = False Then
'    usMsgbox ("Não é permitido alterar " & TextoMsg & " deste " & TextoMsg1 & ", pois o mesmo já está empenhando a produção."), vbExclamation, "CAPRIND v5.0"
'    TBCotacao.Close
'    TBAbrir.Close
'    FunVerifAltCodQtde = False
'    Exit Function
'End If
'TBAbrir.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBCotacao!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido alterar " & TextoMsg & " deste " & TextoMsg1 & ", pois o mesmo já está empenhando o estoque."), vbExclamation, "CAPRIND v5.0"
    TBCotacao.Close
    TBAbrir.Close
    FunVerifAltCodQtde = False
    Exit Function
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function



