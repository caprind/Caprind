VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Prod_Serv_NFSe 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Dados da NFSe"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   1785
   ClientWidth     =   15270
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
   Icon            =   "frmFaturamento_Prod_Serv_NFSe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
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
      Height          =   840
      Left            =   60
      TabIndex        =   35
      Top             =   2430
      Width           =   15135
      Begin VB.TextBox txtRPS 
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
         MaxLength       =   60
         TabIndex        =   43
         ToolTipText     =   "Número do RPS."
         Top             =   390
         Width           =   1095
      End
      Begin VB.ComboBox cmbCertificado 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":1042
         Left            =   7110
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":1044
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Certificado."
         Top             =   390
         Width           =   7845
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Série da nota fiscal."
         Top             =   390
         Width           =   585
      End
      Begin VB.TextBox txtProtocolo 
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
         Left            =   5310
         Locked          =   -1  'True
         MaxLength       =   44
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Número do protocolo."
         Top             =   390
         Width           =   1785
      End
      Begin VB.TextBox txtNota 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Número da nota fiscal."
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Status NFS-e."
         Top             =   390
         Width           =   2355
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número RPS*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10597
         TabIndex        =   40
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   2452
         TabIndex        =   39
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número do protocolo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   5452
         TabIndex        =   38
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1432
         TabIndex        =   37
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3885
         TabIndex        =   36
         Top             =   180
         Width           =   465
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   21
      Top             =   9120
      Width           =   15135
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
         Left            =   9690
         TabIndex        =   14
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
         Left            =   3120
         TabIndex        =   13
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11910
         TabIndex        =   18
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Prod_Serv_NFSe.frx":1046
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
         Left            =   11370
         TabIndex        =   17
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Prod_Serv_NFSe.frx":47EA
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
         Left            =   10260
         TabIndex        =   15
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
         Left            =   10830
         TabIndex        =   16
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Prod_Serv_NFSe.frx":82F3
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
         Left            =   12450
         TabIndex        =   19
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Prod_Serv_NFSe.frx":C3E2
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
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13200
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2430
         TabIndex        =   23
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3750
         TabIndex        =   22
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   60
      TabIndex        =   27
      Top             =   960
      Width           =   15135
      Begin VB.ComboBox cmbCodigoCNAE 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":FC6E
         Left            =   9615
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":FC70
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Código atividade do CNAE."
         Top             =   390
         Width           =   5325
      End
      Begin VB.ComboBox cmbRegimeEspecialTributacao 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":FC72
         Left            =   180
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":FC8B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Regime especial de tributação."
         Top             =   390
         Width           =   4725
      End
      Begin VB.ComboBox cmbExigibilidadeISS 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":FD4E
         Left            =   180
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":FD6A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Exigibilidade do ISS."
         Top             =   990
         Width           =   4725
      End
      Begin VB.ComboBox cmbOperacao 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":FE27
         Left            =   4920
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":FE3D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Operação."
         Top             =   990
         Width           =   3645
      End
      Begin VB.TextBox txtCodTributacao 
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
         Left            =   12105
         MaxLength       =   60
         TabIndex        =   6
         ToolTipText     =   "Código de tributação."
         Top             =   990
         Width           =   2835
      End
      Begin VB.ComboBox cmbNaturezaTributacao 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":FEBC
         Left            =   4920
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":FED5
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Natureza tributação."
         Top             =   390
         Width           =   4680
      End
      Begin VB.ComboBox cmbTipoTributacao 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_NFSe.frx":FF94
         Left            =   8580
         List            =   "frmFaturamento_Prod_Serv_NFSe.frx":FFB0
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Tipo de tributação."
         Top             =   990
         Width           =   3510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código atividade do CNAE"
         Height          =   195
         Index           =   4
         Left            =   11340
         TabIndex        =   41
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Natureza de tributação"
         Height          =   195
         Index           =   2
         Left            =   6390
         TabIndex        =   34
         Top             =   180
         Width           =   1665
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regime especial de tributação"
         Height          =   195
         Left            =   1425
         TabIndex        =   32
         Top             =   180
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exigibilidade do ISS"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   31
         Top             =   780
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação*"
         Height          =   195
         Index           =   1
         Left            =   6345
         TabIndex        =   30
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de tributação do município"
         Height          =   195
         Index           =   6
         Left            =   12315
         TabIndex        =   29
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de tributação"
         Height          =   195
         Index           =   3
         Left            =   9645
         TabIndex        =   28
         Top             =   780
         Width           =   1305
      End
   End
   Begin VB.TextBox txtID_nota 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2400
      TabIndex        =   26
      Text            =   "0"
      Top             =   5010
      Visible         =   0   'False
      Width           =   405
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   9750
      Width           =   15135
      _ExtentX        =   26696
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   33
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   1720
      ButtonCount     =   11
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
      ButtonCaption2  =   "Enviar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Enviar NFS-e para o sefaz (F6)"
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
      ButtonWidth2    =   38
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Cancelar"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Cancelar nota fiscal (F4)"
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
      ButtonLeft3     =   82
      ButtonTop3      =   2
      ButtonWidth3    =   50
      ButtonHeight3   =   21
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
      ButtonLeft4     =   134
      ButtonTop4      =   2
      ButtonWidth4    =   51
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Consultar status"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Consultar status da nota."
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
      ButtonLeft5     =   187
      ButtonTop5      =   2
      ButtonWidth5    =   87
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Log de erros"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Consultar logs de erros do envio."
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
      ButtonLeft6     =   276
      ButtonTop6      =   2
      ButtonWidth6    =   68
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Marcar enviada"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Marcar/desmarcar como enviada para a prefeitura."
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
      ButtonLeft7     =   346
      ButtonTop7      =   2
      ButtonWidth7    =   82
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Marcar autorizada"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Marcar como autorizada."
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
      ButtonLeft8     =   430
      ButtonTop8      =   2
      ButtonWidth8    =   95
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonAlignment9=   2
      ButtonType9     =   1
      ButtonStyle9    =   -1
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   -1
      ButtonLeft9     =   527
      ButtonTop9      =   4
      ButtonWidth9    =   2
      ButtonHeight9   =   54
      ButtonCaption10 =   "Ajuda"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Ajuda (F1)"
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
      ButtonLeft10    =   531
      ButtonTop10     =   2
      ButtonWidth10   =   36
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Sair"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Sair (Esc)"
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
      ButtonLeft11    =   569
      ButtonTop11     =   2
      ButtonWidth11   =   26
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9390
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Prod_Serv_NFSe.frx":1005B
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView ListaNota 
      Height          =   5790
      Left            =   60
      TabIndex        =   12
      Top             =   3285
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   10213
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483641
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
         Text            =   "Empresa"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Série"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   6271
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Enviada"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Status NFe"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "IDempresa"
         Object.Width           =   0
      EndProperty
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_NFSe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XML_ As String
Dim TBLISTA_Faturamento_NFSe As ADODB.Recordset 'OK
Dim NomeArquivo As String
Dim CidadeNFSe As String
Dim reter_PIS As Boolean
Dim reter_Cofins As Boolean
Dim reter_CSLL As Boolean
Dim reter_IR As Boolean
Dim reter_INSS As Boolean

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_NFSe.AbsolutePage <> 2 Then
    If TBLISTA_Faturamento_NFSe.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Faturamento_NFSe.PageCount - 1)
    Else
        TBLISTA_Faturamento_NFSe.AbsolutePage = TBLISTA_Faturamento_NFSe.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Faturamento_NFSe.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Faturamento_NFSe.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Faturamento_NFSe.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_NFSe.AbsolutePage = 1
ProcExibePagina (TBLISTA_Faturamento_NFSe.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_NFSe.AbsolutePage <> -3 Then
    If TBLISTA_Faturamento_NFSe.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Faturamento_NFSe.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Faturamento_NFSe.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_NFSe.AbsolutePage = TBLISTA_Faturamento_NFSe.PageCount
ProcExibePagina (TBLISTA_Faturamento_NFSe.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Initialize()
On Error GoTo tratar_erro

    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
    
ProcCarregaToolBar1 Me, 15135, 11, True

If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Dados da NFSe"
Else
    Caption = "Estoque - Nota fiscal - Dados da NFSe"
End If
ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me

procCarregaCertificado cmbCertificado

'Set spdNFSe = New NFSex.spdNFSeX
'Set spdProxyNFSe = New NFSex.spdProxyNFSeX

ProcCarregaListaNota (1)
ProcLimpaCampos
With frmFaturamento_Prod_Serv
    'Puxa certificado salvo no cadastro da empresa
    If .txtId <> "" And .txtId <> "0" And .txtDtValidacao <> "" Then
        txtID_nota = .txtId
        txtNota = IIf(.txtNFiscal = "", Null, .txtNFiscal)
        txtSerie = .txtSerie
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select CodigoAtividade, Protocolo, Status, RegimeEspecialTributacao, NaturezaTributacao, TipoTributacao, ExigibilidadeISS, Operacao, CodTributacaoMun, ID_nota from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        TBAbrir.Close
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select E.CertificadoDigital, E.CODIGO from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal N ON E.Codigo = N.ID_empresa where N.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = False Then
            If IsNull(TBFI!Certificadodigital) = False Then cmbCertificado = TBFI!Certificadodigital
            
            procCarregaCodigoCNAE TBFI!CODIGO
            CidadeNFSe = FunVerifCidadeEmpresa(TBFI!CODIGO)
            If CidadeNFSe <> "Indaiatuba" Then procConfiguraComponente
        End If
        TBFI.Close
    End If
End With

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o certificado."), vbExclamation, "CAPRIND v5.0"
        TBFI.Close
    Else
        USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    End If
End Sub

Sub ProcCarregaListaNota(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaNota.ListItems.Clear
With frmFaturamento_Prod_Serv
    If .Strsql_FaturamentoNFSe = "" Then Exit Sub
    Set TBLISTA_Faturamento_NFSe = CreateObject("adodb.recordset")
    TBLISTA_Faturamento_NFSe.Open .Strsql_FaturamentoNFSe, Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA_Faturamento_NFSe.EOF = False Then ProcExibePagina (Pagina)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaNota.ListItems.Clear
TBLISTA_Faturamento_NFSe.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Faturamento_NFSe.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_NFSe.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Faturamento_NFSe.RecordCount - IIf(Pagina > 1, (TBLISTA_Faturamento_NFSe.PageSize * (Pagina - 1)), 0), TBLISTA_Faturamento_NFSe.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Faturamento_NFSe.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaNota.ListItems
        .Add , , TBLISTA_Faturamento_NFSe!ID
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Empresa from Empresa where Codigo = " & IIf(IsNull(TBLISTA_Faturamento_NFSe!ID_empresa), 0, TBLISTA_Faturamento_NFSe!ID_empresa), Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        End If
        TBAbrir.Close
        
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Faturamento_NFSe!dt_DataEmissao), "", (Format(TBLISTA_Faturamento_NFSe!dt_DataEmissao, "dd/mm/yy")))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Faturamento_NFSe!int_NotaFiscal), "", TBLISTA_Faturamento_NFSe!int_NotaFiscal)
        If IsNull(TBLISTA_Faturamento_NFSe!TipoNF) = False Then
            If TBLISTA_Faturamento_NFSe!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLISTA_Faturamento_NFSe!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
            If TBLISTA_Faturamento_NFSe!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(4) = TipoNF2
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Faturamento_NFSe!Serie), "", TBLISTA_Faturamento_NFSe!Serie)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Faturamento_NFSe!dbl_Valor_Total_Nota), "0,00", Format(TBLISTA_Faturamento_NFSe!dbl_Valor_Total_Nota, "###,##0.00"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Faturamento_NFSe!txt_Razao_Nome), "", TBLISTA_Faturamento_NFSe!txt_Razao_Nome)
        .Item(.Count).SubItems(8) = IIf(TBLISTA_Faturamento_NFSe!Imprimir = True, "Sim", "Não")
        .Item(.Count).SubItems(9) = IIf(TBLISTA_Faturamento_NFSe!Int_status = 1, "Ativa", "Cancelada")
        .Item(.Count).SubItems(10) = FunVerifStatusNFe(TBLISTA_Faturamento_NFSe!ID)
        .Item(.Count).SubItems(11) = TBLISTA_Faturamento_NFSe!ID_empresa
    End With
    TBLISTA_Faturamento_NFSe.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Faturamento_NFSe.RecordCount
If TBLISTA_Faturamento_NFSe.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Faturamento_NFSe.PageCount
ElseIf TBLISTA_Faturamento_NFSe.AbsolutePage = adPosEOF Then
    lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_NFSe.PageCount & " de: " & TBLISTA_Faturamento_NFSe.PageCount
Else
    lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_NFSe.AbsolutePage - 1 & " de: " & TBLISTA_Faturamento_NFSe.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

cmbRegimeEspecialTributacao.ListIndex = -1
cmbNaturezaTributacao.ListIndex = -1
cmbTipoTributacao.ListIndex = -1
cmbExigibilidadeISS.ListIndex = -1
cmbOperacao.ListIndex = -1
txtCodTributacao = ""
cmbCodigoCNAE.Clear

txtRPS = ""
txtNota = ""
txtSerie = ""
txtStatus = ""
txtProtocolo = ""
txtID_nota = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

With cmbRegimeEspecialTributacao
    Select Case TBAbrir!RegimeEspecialTributacao
        Case "1": .Text = "1 - Microempresa municipal"
        Case "2": .Text = "2 - Estimativa"
        Case "3": .Text = "3 - Sociedade de profissionais"
        Case "4": .Text = "4 - Cooperativa"
        Case "5": .Text = "5 - Microempresário Individual (MEI)"
        Case "6": .Text = "6 - Microempresário e Empresa de Pequeno Porte, (ME EPP)"
    End Select
End With

With cmbNaturezaTributacao
    Select Case TBAbrir!NaturezaTributacao
        Case "1": .Text = "1 - Simples Nacional"
        Case "2": .Text = "2 - Fixo"
        Case "3": .Text = "3 - Depósito em Juízo"
        Case "4": .Text = "4 - Exigibilidade suspensa por decisão judicial"
        Case "5": .Text = "5 - Exigibilidade suspensa por procedimento administrativo"
        Case "6": .Text = "6 - Isenção parcial"
    End Select
End With

With cmbTipoTributacao
    Select Case TBAbrir!TipoTributacao
        Case "1": .Text = "1 - Isenta de ISS"
        Case "2": .Text = "2 - Imune"
        Case "3": .Text = "3 - Não Incidência no Município"
        Case "4": .Text = "4 - Não Tributável"
        Case "5": .Text = "5 - Retida"
        Case "6": .Text = "6 - Tributável dentro do município"
        Case "7": .Text = "7 - Tributável fora do município"
    End Select
End With

With cmbExigibilidadeISS
    Select Case TBAbrir!ExigibilidadeISS
        Case "1": .Text = "1 - Exigível"
        Case "2": .Text = "2 - Não incidência"
        Case "3": .Text = "3 - Isenção"
        Case "4": .Text = "4 - Exportação"
        Case "5": .Text = "5 - Imunidade"
        Case "6": .Text = "6 - Exigibilidade Suspensa por Decisão Judicial"
        Case "7": .Text = "7 - Exigibilidade Suspensa por Processo Administrativo"
    End Select
End With

With cmbOperacao
    Select Case TBAbrir!Operacao
        Case "A": .Text = "A - Sem Dedução"
        Case "B": .Text = "B - Com Dedução/Materiais"
        Case "C": .Text = "C - Imune/Isenta de ISSQN"
        Case "D": .Text = "D - Devolução/Simples Remessa"
        Case "J": .Text = "J – Intermediação"
    End Select
End With

If ListaNota.SelectedItem.SubItems(8) = "Cancelada" Then txtStatus = "Cancelada" Else txtStatus = FunVerifStatusNFe(TBAbrir!ID_nota)
txtCodTributacao = IIf(IsNull(TBAbrir!CodTributacaoMun), "", TBAbrir!CodTributacaoMun)
txtProtocolo = IIf(IsNull(TBAbrir!protocolo), "", TBAbrir!protocolo)

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select RPS from tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    txtRPS = IIf(IsNull(TBFI!RPS), "", TBFI!RPS)
End If
TBFI.Close

'Puxa certificado salvo no cadastro da empresa
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select E.CertificadoDigital, E.Codigo from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal NF ON E.Codigo = NF.ID_empresa where NF.ID = " & TBAbrir!ID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    If IsNull(TBFI!Certificadodigital) = False Then cmbCertificado = TBFI!Certificadodigital
    
    procCarregaCodigoCNAE TBFI!CODIGO
    
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "SELECT codigo, descricao FROM Empresa_CNAE_atividade WHERE ID_empresa = " & TBFI!CODIGO & " AND codigo = " & IIf(IsNull(TBAbrir!CodigoAtividade), 0, TBAbrir!CodigoAtividade), Conexao, adOpenKeyset, adLockReadOnly
    If TBCFOP.EOF = False Then
        cmbCodigoCNAE = TBCFOP!CODIGO & " - " & TBCFOP!Descricao
    End If
    TBCFOP.Close
End If
TBFI.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        TBFI.Close
    Else
        USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    End If
End Sub

Private Sub ListaNota_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaNota
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaNota, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaNota_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaNota
    If .ListItems.Count = 0 Then Exit Sub
    ProcLimpaCampos
    CodigoLista = .SelectedItem.index
    txtID_nota = .SelectedItem
    txtNota = .SelectedItem.SubItems(3)
    txtSerie = .SelectedItem.SubItems(5)
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select CodigoAtividade, Protocolo, Status, RegimeEspecialTributacao, NaturezaTributacao, TipoTributacao, ExigibilidadeISS, Operacao, CodTributacaoMun, ID_nota from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then ProcPuxaDados
    TBAbrir.Close
    
    CidadeNFSe = FunVerifCidadeEmpresa(.SelectedItem.SubItems(11))
    If CidadeNFSe <> "Indaiatuba" Then procConfiguraComponente
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcMarcarDesmarcarEnviada()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaNota
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente marcar/desmarcar como enviada esta(s) nota(s) fiscal(ais)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBAliquota = CreateObject("adodb.recordset")
            TBAliquota.Open "Select Imprimir, int_NotaFiscal, ID, TipoNF, Serie from tbl_Dados_Nota_Fiscal where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAliquota.EOF = False Then
                If TBAliquota!Imprimir = True Then
                    Evento = "Desmarcar"
                    TBAliquota!Imprimir = False
                Else
                    Evento = "Marcar"
                    TBAliquota!Imprimir = True
                End If
                TBAliquota.Update
                '==================================
                Modulo = Formulario
                Evento = Evento & " como enviada"
                ID_documento = .ListItems(InitFor)
                If IsNull(TBAliquota!int_NotaFiscal) = True Or TBAliquota!int_NotaFiscal = "" Then DocumentoTexto1 = "Nº ordem: " & TBAliquota!ID Else DocumentoTexto1 = "Nº nota: " & TBAliquota!int_NotaFiscal
                Documento = DocumentoTexto1 & " - Tipo: " & TBAliquota!TipoNF & " - Série: " & TBAliquota!Serie
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBAliquota.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) nota(s) fiscal(ais) antes de marcar/desmarcar como enviada."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Nota(s) fiscal(ais) marcada(s)/desmarcada(s) como enviada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaNota (1)
    
    With frmFaturamento_Prod_Serv
        .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcMarcarDesmarcarAuto()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtID_nota = 0 Then
    USMsgBox ("Informe a nota antes de marcar como autorizada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus <> "" And txtStatus <> "Rejeição de envio" Then
    USMsgBox ("Não é permitido marcar como autorizada, pois esta nota fiscal já foi enviada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

'Verifica se é NFSe
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota & " and TipoNF = 'M1'", Conexao, adOpenKeyset, adLockReadOnly
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido marcar como autorizada, pois a mesma não é uma nota fiscal de serviço."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * From tbl_Dados_Nota_Fiscal where id = " & txtID_nota & " and DtValidacao IS NULL", Conexao, adOpenKeyset, adLockReadOnly
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido marcar como autorizada, pois esta nota fiscal ainda não foi validada."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If
TBGravar.Close

Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = 0 where id_nota = " & txtID_nota

'==================================
Acao = "Marcar como autorizada"
Modulo = Formulario
ID_documento = txtID_nota
With frmFaturamento_Prod_Serv
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
End With
Documento1 = ""
ProcGravaEvento
'==================================

ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And ListaNota.ListItems.Count <> 0 Then
    ListaNota.SelectedItem = ListaNota.ListItems(CodigoLista)
    ListaNota.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

If cmbRegimeEspecialTributacao <> "" Then TBGravar!RegimeEspecialTributacao = Left(cmbRegimeEspecialTributacao, 1) Else TBGravar!RegimeEspecialTributacao = Null
If cmbNaturezaTributacao <> "" Then TBGravar!NaturezaTributacao = Left(cmbNaturezaTributacao, 1) Else TBGravar!NaturezaTributacao = Null
If cmbTipoTributacao <> "" Then TBGravar!TipoTributacao = Left(cmbTipoTributacao, 1) Else TBGravar!TipoTributacao = Null
If cmbExigibilidadeISS <> "" Then TBGravar!ExigibilidadeISS = Left(cmbExigibilidadeISS, 1) Else TBGravar!ExigibilidadeISS = Null
If cmbOperacao <> "" Then TBGravar!Operacao = Left(cmbOperacao, 1) Else TBGravar!Operacao = Null
TBGravar!CodTributacaoMun = txtCodTributacao
If cmbCodigoCNAE <> "" Then TBGravar!CodigoAtividade = Left(cmbCodigoCNAE, 9) Else TBGravar!CodigoAtividade = Null
Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set RPS = " & txtRPS & " where id = " & txtID_nota

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If txtID_nota = 0 Then
    USMsgBox ("Informe a nota antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus <> "" And txtStatus <> "Rejeição de envio" Then
    USMsgBox ("Não é permitido salvar, pois esta nota fiscal já foi enviada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
'If CidadeNFSe = "Indaiatuba" Then
'    usMsgbox ("Opção não disponivel na cidade de " & CidadeNFSe & "."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If

Acao = "salvar"
If cmbOperacao = "" Then
    NomeCampo = "a operação"
    ProcVerificaAcao
    cmbOperacao.SetFocus
    Exit Sub
End If

If txtRPS = "" Then
    NomeCampo = "o número da RPS"
    ProcVerificaAcao
    txtRPS.SetFocus
    Exit Sub
End If

'Verifica se é NFSe
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota & " and TipoNF = 'M1'", Conexao, adOpenKeyset, adLockReadOnly
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois a mesma não é uma nota fiscal de serviço."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * From tbl_Dados_Nota_Fiscal where id = " & txtID_nota & " and DtValidacao IS NULL", Conexao, adOpenKeyset, adLockReadOnly
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois esta nota fiscal ainda não foi validada."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If
TBGravar.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select RPS FROM tbl_Dados_Nota_Fiscal WHERE RPS = " & txtRPS & " AND ID <> " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois já existe uma nota com este número de RPS."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If
TBGravar.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dados da NFSe"
Else
    TBGravar.AddNew
    USMsgBox ("Novos dados da nota fiscal cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novos dados da NFSe"
    TBGravar!ID_nota = txtID_nota
End If
ProcEnviaDados
TBGravar.Update
TBGravar.Close
'==================================
Modulo = Formulario
ID_documento = txtID_nota
With frmFaturamento_Prod_Serv
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
End With
Documento1 = ""
ProcGravaEvento
'==================================

ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And ListaNota.ListItems.Count <> 0 Then
    ListaNota.SelectedItem = ListaNota.ListItems(CodigoLista)
    ListaNota.SetFocus
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
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: procEnviar
    Case 3: procCancelarNFSe
    Case 4: ProcImprimir
    Case 5: procConsultar
    Case 6: procLogErros
    Case 7: ProcMarcarDesmarcarEnviada
    Case 8: ProcMarcarDesmarcarAuto
    'Case 10: ProcAjuda
    Case 11: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcCriarTX2()
On Error GoTo tratar_erro

Set ArqTXT = GerArqPastas.CreateTextFile(Localrel & "\NFSe\TX2\" & NomeArquivo & ".tx2", True)
With ArqTXT
    'Dados da empresa
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Empresa.* from Empresa INNER JOIN tbl_Dados_Nota_Fiscal ON Empresa.Codigo = tbl_Dados_Nota_Fiscal.ID_empresa where tbl_Dados_Nota_Fiscal.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFI.EOF = False Then
        If IsNull(TBFI!CNPJ) = True Then CPFCNPJ = "99999999999999" Else CPFCNPJ = ReturnNumbersOnly(TBFI!CNPJ)
        IM = ReturnNumbersOnly(IIf(IsNull(TBFI!IM), "", TBFI!IM))
        RazaoSocial = IIf(IsNull(TBFI!Razao), "", TBFI!Razao)
        If IsNull(TBFI!UF) = False And IsNull(TBFI!Cidade) = False Then CodigoCidade = FunVerificaCodMunicipio(TBFI!Cidade, TBFI!UF)
        DescricaoCidadePrestacao = IIf(IsNull(TBFI!Cidade), "", "DescricaoCidadePrestacao=" & TBFI!Cidade)
        If TBFI!Simples = True Then SimplesNacional = "OptanteSimplesNacional=1" Else SimplesNacional = "OptanteSimplesNacional=2"
        If TBFI!Cultural = True Then Cultural = "IncentivadorCultural=1" Else Cultural = "IncentivadorCultural=2"
        If TBFI!Fiscal = True Then IncentivoFiscal = "IncentivoFiscal=1" Else IncentivoFiscal = "IncentivoFiscal=2"
        If cmbCodigoCNAE = "" Then CodigoCNAE = IIf(IsNull(TBFI!CNAE), "", "CodigoCnae=" & FunTamanhoTextoZeroDir(ReturnNumbersOnly(TBFI!CNAE), 9))
    End If
    TBFI.Close
    
    'Dados da nota de serviço
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockReadOnly
    If TBProduto.EOF = False Then
        If IsNull(TBProduto!RegimeEspecialTributacao) = False And TBProduto!RegimeEspecialTributacao <> "" Then RegimeEspecialTributacao = "RegimeEspecialTributacao=" & TBProduto!RegimeEspecialTributacao
        If IsNull(TBProduto!NaturezaTributacao) = False And TBProduto!NaturezaTributacao <> "" Then NaturezaTributacao = "NaturezaTributacao=" & TBProduto!NaturezaTributacao
        If IsNull(TBProduto!TipoTributacao) = False And TBProduto!TipoTributacao <> "" Then TipoTributacao = "TipoTributacao=" & TBProduto!TipoTributacao
        If IsNull(TBProduto!ExigibilidadeISS) = False And TBProduto!ExigibilidadeISS <> "" Then ExigibilidadeISS = "ExigibilidadeISS=" & TBProduto!ExigibilidadeISS
        If IsNull(TBProduto!Operacao) = False And TBProduto!Operacao <> "" Then Operacao = "Operacao=" & TBProduto!Operacao
        If IsNull(TBProduto!CodTributacaoMun) = False And TBProduto!CodTributacaoMun <> "" Then CodigoTributacaoMunicipio = "CodigoTributacaoMunicipio=" & TBProduto!CodTributacaoMun
        If cmbCodigoCNAE <> "" Then CodigoCNAE = IIf(IsNull(TBProduto!CodigoAtividade), "", "CodigoCnae=" & TBProduto!CodigoAtividade)
    End If
    TBProduto.Close
    
    'Dado do serviço
    DescServico = ""
    valor = 0
    IssRetido = ""
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select P.Cod_servico_NFSE, NFP.ISS, NFP.CSLL_Serv, NFP.IRRF_Serv, NFP.INSS_Serv, NFP.Cofins_Serv, NFP.PIS_Serv, P.Cod_servico_NFSE, NFP.txt_Descricao, NFP.PCCliente, NFP.Retencao_ISSQN, NFP.Retencao_PIS, NFP.Retencao_Cofins, NFP.Retencao_CSLL, NFP.Retencao_INSS, NFP.Retencao_IRRF, NFP.vlriss from tbl_Detalhes_Nota NFP INNER JOIN projproduto P ON P.Desenho = NFP.int_Cod_Produto where NFP.ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockReadOnly
    If TBProduto.EOF = False Then
        CodServico = IIf(IsNull(TBProduto!Cod_servico_NFSE), "", "CodigoItemListaServico=" & TBProduto!Cod_servico_NFSE)
        Do While TBProduto.EOF = False
            If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then DescServ = Left(TBProduto!Txt_descricao, 80) & " - Ped. " & Trim(Left(TBProduto!PCCliente, 12)) Else DescServ = Left(TBProduto!Txt_descricao, 90)
            If DescServico = "" Then DescServico = DescServ Else DescServico = DescServico & "|" & DescServ
            
            If TBProduto!Retencao_ISSQN = True Then
                valor = valor + TBProduto!VlrISS
                IssRetido = "IssRetido=true"
            Else
                'Se passar pelo retido true não pode mudar pois false é NÃO RETIDO
                If IssRetido = "" Then IssRetido = "IssRetido=false"
            End If
            
            'Verifica se vai reter imposto
            If IsNull(TBProduto!Retencao_PIS) = True Or TBProduto!Retencao_PIS = False Then
                reter_PIS = False
                AliquotaPIS = "AliquotaPIS=0.00"
            Else
                reter_PIS = True
                If IsNull(TBProduto!PIS_Serv) = False Then AliquotaPIS = "AliquotaPIS=" & Format(TBProduto!PIS_Serv, "0.00")
            End If
            
            If IsNull(TBProduto!Retencao_Cofins) = True Or TBProduto!Retencao_Cofins = False Then
                reter_Cofins = False
                AliquotaCofins = "AliquotaCOFINS=0.00"
            Else
                reter_Cofins = True
                If IsNull(TBProduto!Cofins_Serv) = False Then AliquotaCofins = "AliquotaCOFINS=" & Format(TBProduto!Cofins_Serv, "0.00")
            End If
            
            If IsNull(TBProduto!Retencao_CSLL) = True Or TBProduto!Retencao_CSLL = False Then
                reter_CSLL = False
                AliquotaCSLL = "AliquotaCSLL=0.00"
            Else
                reter_CSLL = True
                If IsNull(TBProduto!CSLL_Serv) = False Then AliquotaCSLL = "AliquotaCSLL=" & Format(TBProduto!CSLL_Serv, "0.00")
            End If
            
            If IsNull(TBProduto!Retencao_INSS) = True Or TBProduto!Retencao_INSS = False Then
                reter_INSS = False
                AliquotaINSS = "AliquotaINSS=0.00"
            Else
                reter_INSS = True
                If IsNull(TBProduto!INSS_Serv) = False Then AliquotaINSS = "AliquotaINSS=" & Format(TBProduto!INSS_Serv, "0.00")
            End If
            
            If IsNull(TBProduto!Retencao_IRRF) = True Or TBProduto!Retencao_IRRF = False Then
                reter_IR = False
                AliquotaIR = "AliquotaIR=0.00"
            Else
                reter_IR = True
                If IsNull(TBProduto!IRRF_Serv) = False Then AliquotaIR = "AliquotaIR=" & Format(TBProduto!IRRF_Serv, "0.00")
            End If
            
            If CidadeNFSe = "Salto" Then
                'Apenas salva esse campo se for simples nacional em salto
                If SimplesNacional = "OptanteSimplesNacional=1" And IsNull(TBProduto!ISS) = False Then AliquotaISS = "AliquotaISS=" & Format(TBProduto!ISS, "0.00")
            Else
                If IsNull(TBProduto!ISS) = False Then AliquotaISS = "AliquotaISS=" & Format(TBProduto!ISS, "0.00")
            End If
            TBProduto.MoveNext
        Loop
    End If
    TBProduto.Close
    
    'Dados dos totais da nota
    DescServico = Left(DescServico, 1950)
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open "Select DAS, dbl_valor_total_iss, Total_CSLL_serv, Total_IRRF_serv, Total_INSS_serv, Total_Cofins_serv, Total_PIS_serv, Valor_total_aprox_tributos, dbl_Valor_Total_Nota_Serv from tbl_Totais_Nota where ID_Nota = " & TBproducao!ID, Conexao, adOpenKeyset, adLockReadOnly
    If TBTotaisnota.EOF = False Then
        If DescServico = "" Then DescServico = "Valor tributos aprox.: " & Format(TBTotaisnota!Valor_total_aprox_tributos, "###,##0.00") Else DescServico = DescServico & "|Valor tributos aprox.: " & Format(TBTotaisnota!Valor_total_aprox_tributos, "###,##0.00")
        DescServico = Left(DescServico, 2000)
        VlrTotalServico = Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "0.00")
        If reter_PIS = True Then ValorPIS = "ValorPIS=" & Format(TBTotaisnota!Total_PIS_serv, "0.00") Else ValorPIS = "ValorPIS=0.00"
        If reter_Cofins = True Then ValorCOFINS = "ValorCOFINS=" & Format(TBTotaisnota!Total_Cofins_serv, "0.00") Else ValorCOFINS = "ValorCOFINS=0.00"
        If reter_INSS = True Then ValorINSS = "ValorINSS=" & Format(TBTotaisnota!Total_INSS_serv, "0.00") Else ValorINSS = "ValorINSS=0.00"
        If reter_IR = True Then ValorIR = "ValorIR=" & Format(TBTotaisnota!Total_IRRF_serv, "0.00") Else ValorIR = "ValorIR=0.00"
        If reter_CSLL = True Then ValorCSLL = "ValorCSLL=" & Format(TBTotaisnota!Total_CSLL_serv, "0.00") Else ValorCSLL = "ValorCSLL=0.00"
        valorISS = "ValorISS=" & Format(TBTotaisnota!dbl_valor_total_iss, "0.00")
        ValorLiquidoNfse = "ValorLiquidoNfse=" & Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv - IIf(reter_PIS = True, TBTotaisnota!Total_PIS_serv, 0) - IIf(reter_Cofins = True, TBTotaisnota!Total_Cofins_serv, 0) - IIf(reter_INSS = True, TBTotaisnota!Total_INSS_serv, 0) - IIf(reter_IR = True, TBTotaisnota!Total_IRRF_serv, 0) - IIf(reter_CSLL = True, TBTotaisnota!Total_CSLL_serv, 0) - valor, "0.00")
    End If
    TBTotaisnota.Close
    
    'Dados do tomador da tabela de cadastro
    CPFCNPJTomador = ""
    EmailTomador = ""
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select idTipoEmpresa, Pais, Tipo_bairro, Bairro, Tipo_endereco, Endereco, Complemento, Email from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBClientes.EOF = False Then
        If TBClientes!idTipoEmpresa = 0 Then
            If Left(TBproducao!txt_tipocliente, 1) = "J" Then CPFCNPJTomador = "99999999999999"
        End If
        TipoBairro = "TipoBairroTomador=" & TBClientes!Tipo_bairro
        BairroTomador = "BairroTomador=" & Left(TBClientes!Bairro, 50)
        TipoEndereco = "TipoLogradouroTomador=" & TBClientes!Tipo_endereco
        EnderecoTomador = "EnderecoTomador=" & Left(TBClientes!Endereco, 50)
        ComplementoTomador = IIf(IsNull(TBClientes!complemento), "", "ComplementoTomador=" & TBClientes!complemento)
        EmailTomador = IIf(IsNull(TBClientes!Email), "", "EmailTomador=" & Left(TBClientes!Email, 60))
        
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select Email from Clientes_Contatos where IDCliente = " & TBproducao!Id_Int_Cliente & " and Enviar_NFe = 'True'", Conexao, adOpenKeyset, adLockReadOnly
        If TBContas.EOF = False Then
            EmailTomador = IIf(IsNull(TBContas!Email), "", "EmailTomador=" & Left(Trim(TBContas!Email), 60))
        End If
        TBContas.Close
    Else
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select idTipoEmpresa, Pais, Tipo_bairro, Bairro, Tipo_endereco, Endereco, Complemento, Email from Compras_fornecedores where IDCliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            If TBClientes!idTipoEmpresa = 0 Then
                If Left(TBproducao!txt_tipocliente, 1) = "J" Then CPFCNPJTomador = "99999999999999"
            End If
            TipoBairro = "TipoBairroTomador=" & TBClientes!Tipo_bairro
            BairroTomador = "BairroTomador=" & Left(TBClientes!Bairro, 50)
            TipoEndereco = "TipoLogradouroTomador=" & TBClientes!Tipo_endereco
            EnderecoTomador = "EnderecoTomador=" & Left(TBClientes!Endereco, 50)
            ComplementoTomador = IIf(IsNull(TBClientes!complemento), "", "ComplementoTomador=" & TBClientes!complemento)
            EmailTomador = IIf(IsNull(TBClientes!Email), "", "EmailTomador=" & Left(TBClientes!Email, 60))
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBproducao!Id_Int_Cliente & " and Enviar_NFe = 'True'", Conexao, adOpenKeyset, adLockReadOnly
            If TBContas.EOF = False Then
                EmailTomador = IIf(IsNull(TBContas!Email), "", "EmailTomador=" & Left(TBContas!Email, 60))
            End If
            TBContas.Close
        End If
    End If
    TBClientes.Close
        
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open "Select mem_DadosAdicionais from tbl_dadosadicionais where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBTotaisnota.EOF = False Then
        If DescServico = "" Then DescServico = Trim(TBTotaisnota!mem_DadosAdicionais) Else DescServico = DescServico & "||" & TBTotaisnota!mem_DadosAdicionais
    End If
    TBTotaisnota.Close
    
    'Dados do tomador da tabela de nota
    If CPFCNPJTomador = "" Then CPFCNPJTomador = "CpfCnpjTomador=" & ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF)
    RazaoSocialTomador = "RazaoSocialTomador=" & Left(TBproducao!txt_Razao_Nome, 100)
    IeTomador = "InscricaoEstadualTomador=" & IIf(IsNull(TBproducao!txt_IE_Cliente), "", ReturnNumbersOnly(TBproducao!txt_IE_Cliente))
    If IsNull(TBproducao!Numero) = True And TBproducao!Numero = "" Then NumeroTomador = "NumeroTomador=" & 0 Else NumeroTomador = "NumeroTomador=" & Left(ReturnNumbersOnly(TBproducao!Numero), 6)
    If IsNull(TBproducao!txt_Municipio) = False And IsNull(TBproducao!txt_UF) = False Then CodigoCidadeTomador = "CodigoCidadeTomador=" & FunVerificaCodMunicipio(TBproducao!txt_Municipio, TBproducao!txt_UF)
    CidadeTomador = "DescricaoCidadeTomador=" & Left(Replace(TBproducao!txt_Municipio, "do Oeste", "d'Oeste"), 50)
    UFTomador = "UfTomador=" & TBproducao!txt_UF
    If TipoEndereco = "" Then TipoEnderecoTomador = "TipoLogradouroTomador=" & "RUA"
    CEPTomador = "CepTomador=" & Left(ReturnNumbersOnly(TBproducao!Txt_CEP), 8)
    If IsNull(TBproducao!txt_Fone_Fax) = False And TBproducao!txt_Fone_Fax <> "" Then TelefoneTomador = "TelefoneTomador=" & Left(ReturnNumbersOnly(TBproducao!txt_Fone_Fax), 10) Else TelefoneTomador = ""
    
    'Dados da Nota
    VlrISSRetido = "ValorISSRetido=" & Format(valor, "0.00")
    IDRps = "IdRps=" & TBproducao!ID
    DataRPS = Format(TBproducao!dt_DataEmissao, "yyyy-mm-dd")
    
    'Inicio
    .WriteLine "formato=tx2"
    .WriteLine "padrao=TecnoNFSe"
    .WriteLine ""
    
    'Incluir
    .WriteLine "INCLUIR"
    .WriteLine "NumeroLote=" & txtRPS
    .WriteLine "QuantidadeRPS=1"
    .WriteLine "Transacao=true"
    .WriteLine "MetodoEnvio=WS"
    .WriteLine "CpfCnpjRemetente=" & CPFCNPJ
    .WriteLine "InscricaoMunicipalRemetente=" & IM
    .WriteLine "RazaoSocialRemetente=" & RazaoSocial
    .WriteLine "CodigoCidadeRemetente=" & CodigoCidade
    .WriteLine "DataInicio=" & DataRPS
    .WriteLine "DataFim=" & DataRPS
    .WriteLine "ValorTotalServicos=" & VlrTotalServico
    .WriteLine "ValorTotalDeducoes=0,00"
    .WriteLine "ValorTotalBaseCalculo=" & VlrTotalServico
    .WriteLine "SALVAR"
    .WriteLine ""
    
    'Prestador
    .WriteLine "INCLUIRRPS"
    .WriteLine IDRps
    .WriteLine "SituacaoNota=1"
    .WriteLine "TipoRps=1"
    .WriteLine "SerieRps=" & TBproducao!Serie
    .WriteLine "NumeroRps=" & txtRPS
    .WriteLine "DataEmissao=" & Format(TBproducao!dt_DataEmissao, "yyyy-mm-dd") & "T" & Left(Format(TBproducao!Hora_emissao, "h:m:s"), 8)
    .WriteLine "Competencia=" & Format(TBproducao!dt_Saida_Entrada, "yyyy-mm-dd")
    .WriteLine "CpfCnpjPrestador=" & CPFCNPJ
    .WriteLine "InscricaoMunicipalPrestador=" & IM
    .WriteLine "RazaoSocialPrestador=" & RazaoSocial
    .WriteLine "CodigoCidadePrestacao=" & CodigoCidade
    .WriteLine DescricaoCidadePrestacao
    
    .WriteLine ""
    
    .WriteLine "DiscriminacaoServico=" & Replace(DescServico, vbCrLf, "|")
    .WriteLine SimplesNacional
    .WriteLine Cultural
    .WriteLine RegimeEspecialTributacao
    .WriteLine NaturezaTributacao
    .WriteLine IncentivoFiscal
    .WriteLine ""
    
    'Tomador
    .WriteLine CPFCNPJTomador
    .WriteLine RazaoSocialTomador
    .WriteLine IeTomador
    .WriteLine TipoEndereco
    .WriteLine EnderecoTomador
    .WriteLine NumeroTomador
    .WriteLine ComplementoTomador
    .WriteLine TipoBairro
    .WriteLine BairroTomador
    .WriteLine CodigoCidadeTomador
    .WriteLine CidadeTomador
    .WriteLine UFTomador
    .WriteLine CEPTomador
    .WriteLine "PaisTomador=1058"
    .WriteLine TelefoneTomador
    .WriteLine EmailTomador
    .WriteLine ""
    
    'Serviço
    .WriteLine CodServico
    .WriteLine CodigoTributacaoMunicipio
    .WriteLine CodigoCNAE
    .WriteLine TipoTributacao
    .WriteLine ExigibilidadeISS
    .WriteLine Operacao
    .WriteLine "MunicipioIncidencia=" & CodigoCidade
    .WriteLine ""
    
    'Valores
    .WriteLine "ValorServicos=" & VlrTotalServico
    .WriteLine AliquotaPIS
    .WriteLine AliquotaCofins
    .WriteLine AliquotaINSS
    .WriteLine AliquotaIR
    .WriteLine AliquotaCSLL
    .WriteLine ValorPIS
    .WriteLine ValorCOFINS
    .WriteLine ValorINSS
    .WriteLine ValorIR
    .WriteLine ValorCSLL
    .WriteLine "OutrasRetencoes=0,00"
    .WriteLine "DescontoIncondicionado=0,00"
    .WriteLine "DescontoCondicionado=0,00"
    .WriteLine "ValorDeducoes=0,00"
    .WriteLine "BaseCalculo=" & VlrTotalServico
    .WriteLine AliquotaISS
    .WriteLine valorISS
    .WriteLine IssRetido
    .WriteLine VlrISSRetido
    .WriteLine ValorLiquidoNfse
    .WriteLine "SALVARRPS"
    .Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procEnviar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

'Verifica o status
If txtStatus = "Autorizada" Or txtStatus = "Cancelada" Then
    USMsgBox ("Não é permitido enviar, pois a mesma já foi enviada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If cmbCertificado = "" Then
    USMsgBox "Informe o certificado antes de enviar.", vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If CidadeNFSe = "Indaiatuba" Then
    USMsgBox ("Opção não disponivel na cidade de " & CidadeNFSe & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

'Verifica se é da cidade de Indaiatuba quando for NFSe
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select E.cidade from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal N ON E.Codigo = N.ID_empresa where N.ID = " & txtID_nota & " AND (E.Cidade = 'Indaiatuba' OR E.Cidade = 'INDAIATUBA')", Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    USMsgBox ("Não é permitido enviar nota para emitentes da cidade de Indaiatuba."), vbExclamation, "CAPRIND v5.0"
    TBFI.Close
    Exit Sub
End If
TBFI.Close

If USMsgBox("Deseja realmente liberar esta(s) nota(s) fiscal(ais)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBproducao.EOF = False Then
        NomeArquivo = Format(Date, "ddmmyyyy") & TBproducao!ID
        ProcCriarTX2
        procEnviaXML
    End If
    TBproducao.Close
    
    ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    With frmFaturamento_Prod_Serv
        .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procEnviaXML()
On Error GoTo tratar_erro
Dim ArquivoLog, XML  As String
'Dim spdNFSeConverter As NFSeConverterX.spdNFSeConverterX
'Set spdNFSeConverter = New NFSeConverterX.spdNFSeConverterX


XML_ = spdNFSeConverter.ConverterEnvioNFSe(Localrel & "\NFSe\TX2\" & NomeArquivo & ".tx2", "")
XML_ = spdProxyNFSe.Assinar(XML_)

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Empresa.* from Empresa INNER JOIN tbl_Dados_Nota_Fiscal ON Empresa.Codigo = tbl_Dados_Nota_Fiscal.ID_empresa where tbl_Dados_Nota_Fiscal.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBFI.EOF = False Then
    If CidadeNFSe = "Salto" Then
        senhaNFSe = "Senha=" & TBFI!SenhaPref
        XML_ = spdProxyNFSe.EnviarSincrono(XML_, senhaNFSe)
    Else
        protocolo = spdProxyNFSe.Enviar(XML_, "")
        txtProtocolo = protocolo
    End If
End If
TBFI.Close

If CidadeNFSe = "Salto" Then
    procTratarRetornoEnvioSincrono (XML_)
Else
    ArquivoLog = spdProxyNFSe.ComponenteNFSe.UltimoLogRecibo
    Set ArqTXT = GerArqPastas.OpenTextFile(ArquivoLog)
    
    XML = ArqTXT.ReadAll
    procTratarRetornoEnvio (XML)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procConfiguraComponente()
On Error GoTo tratar_erro
''Dim spdNFSeConverter As NFSeConverterX.spdNFSeConverterX
''Set spdNFSeConverter = New NFSeConverterX.spdNFSeConverterX
'
''Dados da empresa
'Set TBCFOP = CreateObject("adodb.recordset")
'TBCFOP.Open "Select E.* from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal NF ON E.Codigo = NF.ID_empresa where NF.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
'If TBCFOP.EOF = False Then
'    If TBCFOP!Cidade = "Salto" Then
'        spdNFSe.Cidade = TBCFOP!Cidade & "SP"
'    Else
'        spdNFSe.Cidade = TBCFOP!Cidade
'    End If
'    spdNFSe.NomeCertificado = cmbCertificado
'    If IsNull(TBCFOP!Senha) = False Then spdNFSe.SenhaCertificado = TBCFOP!Senha
'    spdNFSe.Proxy = 0
'    spdNFSe.CNPJ = ReturnNumbersOnly(TBCFOP!CNPJ)
'    spdNFSe.InscricaoMunicipal = ReturnNumbersOnly(TBCFOP!IM)
'    spdNFSe.Ambiente = akProducao
'    spdNFSe.ArquivoLocais = Localrel & "\NFSe\" & "Arquivos\nfseLocais.ini"
'    spdNFSe.ArquivoServidoresHom = Localrel & "\NFSe\" & "Arquivos\nfseServidoresHom.ini"
'    spdNFSe.ArquivoServidoresProd = Localrel & "\NFSe\" & "Arquivos\nfseServidoresProd.ini"
'    spdNFSe.DiretorioEsquemas = Localrel & "\NFSe\" & "Arquivos\Esquemas"
'    spdNFSe.DiretorioTemplates = Localrel & "\NFSe\" & "Arquivos\Templates"
'    spdNFSe.DiretorioXmlImpressao = Localrel & "\NFSe\" & "Impressao"
'
'    spdNFSe.DiretorioLog = Localrel & "\NFSe\Log"
'    spdNFSe.MappingFileName = "Mapping.txt"
'    spdNFSe.DiretorioLogErro = Localrel & "\NFSe\LogErro"
'
'    spdNFSeConverter.DiretorioEsquemas = spdNFSe.DiretorioEsquemas
'    spdNFSeConverter.DiretorioScripts = spdNFSe.DiretorioEsquemas + "\..\Scripts\"
'    spdNFSeConverter.Cidade = spdNFSe.Cidade
'
'    spdProxyNFSe.ComponenteNFSe = spdNFSe
'End If
'TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procTratarRetornoEnvio(aXML As String)
'On Error GoTo tratar_erro
'Dim Ret_ As New spdRetEnvioNFSe
'
'Set Ret_ = spdNFSeConverter.ConverterRetEnvioNFSeTipo(aXML)
'Set TBGravar = CreateObject("adodb.recordset")
'TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
'If TBGravar.EOF = False Then
'    TBGravar!protocolo = Ret_.NumeroProtocolo
'
'    If Ret_.motivo <> "" Then
'        TBGravar!LogErro = "Motivo erro da nota " & txtNota & ": " & Ret_.motivo
'        TBGravar!status = 2
'        If Ret_.NumeroProtocolo = "" Then
'            USMsgBox ("Motivo erro da nota " & txtNota & ": " & Ret_.motivo), vbExclamation, "CAPRIND v5.0"
'        End If
'    Else
'        TBGravar!LogErro = Null
'    End If
'
'    TBGravar.Update
'End If
'TBGravar.Close
'
'If Ret_.NumeroProtocolo <> "" Then
'    XML_ = spdProxyNFSe.ConsultarLote(Ret_.NumeroProtocolo)
'    procRetornoLote (XML_)
'End If
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procTratarRetornoEnvioSincrono(aXML As String)
On Error GoTo tratar_erro
    
'Set TBGravar = CreateObject("adodb.recordset")
'TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
'If TBGravar.EOF = False Then
'    If spdNFSeConverter.ConverterRetEnvioSincronoNFSeTipo(aXML).status = 1 Then
'        'Processando
'        TBGravar!status = 1
'        USMsgBox ("Nota fiscal processando."), vbInformation, "CAPRIND v5.0"
'    Else
'        If spdNFSeConverter.ConverterRetEnvioSincronoNFSeTipo(aXML).status = 2 Then
'            'Erro
'            TBGravar!status = 2
'            TBGravar!LogErro = "Motivo erro da nota " & txtNota & ": " & spdNFSeConverter.ConverterRetEnvioSincronoNFSeTipo(aXML).motivo
'            USMsgBox ("Erro ao enviar a nota, verifique o XML de retorno."), vbExclamation, "CAPRIND v5.0"
'        Else
'            'Sucesso
'            TBGravar!status = 0
'            USMsgBox ("Nota fiscal aprovada com sucesso."), vbInformation, "CAPRIND v5.0"
'        End If
'    End If
'    TBGravar!CaminhoLog = spdProxyNFSe.ComponenteNFSe.UltimoLogRecibo
'
'    TBGravar.Update
'End If
'TBGravar.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procRetornoLote(aXML As String)
On Error GoTo tratar_erro

'Set TBGravar = CreateObject("adodb.recordset")
'TBGravar.Open "Select Status, LogErro FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
'If TBGravar.EOF = False Then
'    If spdNFSeConverter.ConverterRetConsultarLoteNFSeTipo(aXML).status = 1 Then
'        'Processando
'        TBGravar!status = 1
'        USMsgBox ("Nota fiscal processando."), vbInformation, "CAPRIND v5.0"
'    ElseIf spdNFSeConverter.ConverterRetConsultarLoteNFSeTipo(aXML).status = 2 Then
'        'Erro
'        TBGravar!status = 2
'        TBGravar!LogErro = "Motivo: " + spdNFSeConverter.ConverterRetConsultarLoteNFSeTipo(aXML).motivo
'        USMsgBox ("Motivo erro da nota " & txtNota & ": " & spdNFSeConverter.ConverterRetConsultarLoteNFSeTipo(aXML).motivo), vbExclamation, "CAPRIND v5.0"
'    Else
'        'Sucesso
'        TBGravar!status = 0
'        USMsgBox ("Nota fiscal aprovada com sucesso."), vbInformation, "CAPRIND v5.0"
'    End If
'
'    TBGravar.Update
'End If
'TBGravar.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcImprimir()
On Error GoTo tratar_erro

'If txtID_nota = 0 Then
'    USMsgBox ("Informe a nota fiscal antes de consultar o status."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'If CidadeNFSe = "Indaiatuba" Then
'    USMsgBox ("Opção não disponivel na cidade de " & CidadeNFSe & "."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'If txtStatus <> "Autorizada" And txtStatus <> "Cancelada" Then
'    USMsgBox ("Só é possível visualizar impressão de notas autorizadas ou canceladas."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'If CidadeNFSe = "Salto" Then
'    Set TBHistProc = CreateObject("adodb.recordset")
'    TBHistProc.Open "Select CaminhoLog FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND CaminhoLog IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
'    If TBHistProc.EOF = False Then
'        Set ArqTXT = GerArqPastas.OpenTextFile(TBHistProc!CaminhoLog)
'        XML_ = ArqTXT.ReadAll
'    Else
'        USMsgBox "Não é possivel imprimir pois não foi encontrado o retorno da prefeitura.", vbExclamation, "CAPRIND v5.0"
'        TBHistProc.Close
'        Exit Sub
'    End If
'    TBHistProc.Close
'Else
'    NumeroNota = CInt(txtNota)
'    XML_ = spdProxyNFSe.ConsultarNota(NumeroNota, "")
'    NomeArquivo = ""
'End If
'
'spdProxyNFSe.ComponenteNFSe.ImpressaoModo = printNFSe
'ConfigurarImpressao XML_
'spdProxyNFSe.ComponenteNFSe.Impressao_VisualizarDocumentoCustom "", ""
    
Exit Sub
tratar_erro:
    If Err.Number = 53 Or Err.Number = 76 Then
        USMsgBox "Não é possivel imprimir pois não foi encontrado o retorno da prefeitura.", vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procCarregaCodigoCNAE(IDEmpresaCNAE As Integer)
On Error GoTo tratar_erro

cmbCodigoCNAE.Clear
cmbCodigoCNAE.AddItem ""
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select codigo, descricao from Empresa_CNAE_atividade where ID_empresa = " & IDEmpresaCNAE, Conexao, adOpenKeyset, adLockReadOnly
Do While TBCFOP.EOF = False
    cmbCodigoCNAE.AddItem TBCFOP!CODIGO & " - " & TBCFOP!Descricao
    TBCFOP.MoveNext
Loop
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procCancelarNFSe()
'On Error GoTo tratar_erro
'Dim Ret_ As New spdRetCancelaNFSe

'If txtID_nota = 0 Then
'    USMsgBox ("Informe a nota fiscal antes de cancelar."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'If CidadeNFSe = "Salto" Or CidadeNFSe = "Indaiatuba" Then
'    USMsgBox ("Opção não disponivel na cidade de " & CidadeNFSe & "."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'If txtStatus <> "Autorizada" Then
'    USMsgBox ("Só é possível cancelar notas aprovadas."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'NumeroNota = CInt(txtNota)
'XML_ = spdProxyNFSe.CancelarNota(NumeroNota)
'Set Ret_ = spdNFSeConverter.ConverterRetCancelarNFSeTipo(XML_)
'Set TBHistProc = CreateObject("adodb.recordset")
'TBHistProc.Open "Select Status, LogErro FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
'If TBHistProc.EOF = False Then
'    Select Case Ret_.status
'        Case 0
'            'Cancelado com sucesso
'            TBHistProc!status = 3
'            Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set int_status = 2 where id = " & txtID_nota
'            USMsgBox ("Nota fiscal cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
'        Case 1
'            'Processando
'            USMsgBox ("Nota fiscal ainda não foi cancelada, em processo de cancelamento."), vbExclamation, "CAPRIND v5.0"
'        Case 2
'            'Erro
'            TBHistProc!LogErro = "Motivo: " + Ret_.motivo
'            USMsgBox ("Não foi possível cancelar, motivo: " + Ret_.motivo), vbExclamation, "CAPRIND v5.0"
'    End Select
'    TBHistProc.Update
'    txtStatus = FunVerifStatusNFe(txtID_nota)
'End If
'TBHistProc.Close
'
'ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
'With frmFaturamento_Prod_Serv
'    If txtStatus = "Cancelada" Then
'        Set TBHistProc = CreateObject("adodb.recordset")
'        TBHistProc.Open "SELECT txt_Tipocliente FROM tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
'        If TBHistProc.EOF = False Then
'            .ProcExcluirArquivosRemessa txtID_nota
'            .ProcExcluirContas txtID_nota, True, IIf(IsNull(TBHistProc!txt_tipocliente), "", TBHistProc!txt_tipocliente)
'            Conexao.Execute "DELETE from ECEV from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = ECEV.ID_faturamento where NFP.ID_nota = " & txtID_nota
'        End If
'        TBHistProc.Close
'    End If
'    .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
'End With
'If CodigoLista <> 0 And ListaNota.ListItems.Count <> 0 Then
'    ListaNota.SelectedItem = ListaNota.ListItems(CodigoLista)
'    ListaNota.SetFocus
'End If
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcExcluirContas(ID_nota As Long, Saida As Boolean, Entrada As Boolean, TipoCliente As String)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & ID_nota & " and CFOP.Devolucao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    'Fornecedor
    If Saida = True And Len(TipoCliente) = 1 Then GoTo ExcluirPagar
    'Cliente
    If Entrada = True And Len(TipoCliente) = 2 Then GoTo ExcluirReceber
Else
    If Saida = True Then GoTo ExcluirReceber Else GoTo ExcluirPagar
End If
TBAbrir.Close

ExcluirReceber:
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_contas_Receber where id_nota = " & ID_nota & " and Bloqueado = 'False' and Status <> 'DUPLICATA DESCONTADA RECOMPRADA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        Do While TBContas.EOF = False
            Conexao.Execute "DELETE from CC_realizado where Operacao = 'Crédito' and ID_financeiro = " & TBContas!IDintconta
            
            If (IsNull(TBContas!Proposta) = True Or TBContas!Proposta = "") And TBContas!Logsit = "N" Then
                'Contas contabeis
                Conexao.Execute "DELETE FROM Familia_financeiro WHERE IDConta = " & TBContas!IDintconta & " and Tipoconta = 'R' and Deposito_transf = 'False'"
                'Fluxo de Caixa
                Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
                'Número dos boletos
                Conexao.Execute "DELETE from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & TBContas!IDintconta
                'Conta
                Conexao.Execute "DELETE from tbl_contas_Receber where IDintconta = " & TBContas!IDintconta
            ElseIf IsNull(TBContas!Proposta) = False And TBContas!Proposta <> "" Then
                    TBContas!ID_nota = 0
                    TBContas!NFiscal = ""
                    TBContas.Update
            End If
            TBContas.MoveNext
        Loop
    End If
    TBContas.Close
    GoTo Prosseguir

ExcluirPagar:
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_ContasPagar where id_nota = " & ID_nota & " and Bloqueado = 'False' and Despesas_NF = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        Do While TBContas.EOF = False
            If (IsNull(TBContas!Txt_pedido) = True Or TBContas!Txt_pedido = "") And TBContas!Logsit = "N" Then
                'Contas contabeis
                Conexao.Execute "DELETE FROM Familia_financeiro WHERE IDConta = " & TBContas!IDintconta & " and Tipoconta = 'P' and Deposito_transf = 'False'"
                'Fluxo de Caixa
                Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
                'Conta
                Conexao.Execute "DELETE from tbl_ContasPagar where IDintconta = " & TBContas!IDintconta
            ElseIf IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" Then
                    TBContas!ID_nota = 0
                    TBContas!txt_ndocumento = ""
                    TBContas.Update
            End If
            TBContas.MoveNext
        Loop
    End If
    TBContas.Close

Prosseguir:
    Conexao.Execute "Update CC set CC.ID_Financeiro = 0 from CC_realizado CC INNER JOIN tbl_Detalhes_Recebimento TBL on CC.ID_duplicata = TBL.ID where TBL.ID_nota = " & ID_nota

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procConsultar()
On Error GoTo tratar_erro
'Dim Ret_ As New spdRetConsultaNFSe
'
'If txtID_nota = 0 Then
'    USMsgBox ("Informe a nota fiscal antes de consultar o status."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'If CidadeNFSe = "Salto" Or CidadeNFSe = "Indaiatuba" Then
'    USMsgBox ("Opção não disponivel na cidade de " & CidadeNFSe & "."), vbExclamation, "CAPRIND v5.0"
'    Exit Sub
'End If
'
'NumeroNota = CInt(txtNota)
'XML_ = spdProxyNFSe.ConsultarNota(NumeroNota, "")
'Set Ret_ = spdNFSeConverter.ConverterRetConsultarNFSeTipo(XML_)
'
'Set TBHistProc = CreateObject("adodb.recordset")
'TBHistProc.Open "Select Status, LogErro FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
'If TBHistProc.EOF = False Then
'    Select Case Ret_.status
'        Case 0
'            'Sucesso
'            If Ret_.situacao = "CANCELADA" Then
'                TBHistProc!status = 3
'                Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set int_status = 2 where id = " & txtID_nota
'            Else
'                TBHistProc!status = 0
'                Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set int_status = 1 where id = " & txtID_nota
'            End If
'        Case 1
'            'Processando
'            TBHistProc!status = 1
'        Case 2
'            'Erro
'            TBHistProc!status = 2
'            TBHistProc!LogErro = "Motivo: " + Ret_.motivo
'    End Select
'    TBHistProc.Update
'
'    txtStatus = FunVerifStatusNFe(txtID_nota)
'End If
'TBHistProc.Close
'USMsgBox ("Nota fiscal com status: " & txtStatus & "."), vbInformation, "CAPRIND v5.0"
'
'ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
'With frmFaturamento_Prod_Serv
'    .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
'End With
'If CodigoLista <> 0 And ListaNota.ListItems.Count <> 0 Then
'    ListaNota.SelectedItem = ListaNota.ListItems(CodigoLista)
'    ListaNota.SetFocus
'End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procLogErros()
On Error GoTo tratar_erro

If txtID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de consultar log de erros."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Sit_REG = 1
frmFaturamento_Prod_Serv_NFSe_Log.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: procCancelarNFSe
    Case vbKeyF5: ProcImprimir
    Case vbKeyF6: procEnviar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ConfigurarImpressao(aXML As String)
On Error GoTo tratar_erro

'Set TBFI = CreateObject("adodb.recordset")
'TBFI.Open "Select E.*, NF.dt_DataEmissao, NF.Hora_emissao from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal NF ON E.Codigo = NF.ID_empresa where NF.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
'If TBFI.EOF = False Then
'    If CidadeNFSe <> "Salto" Then
'        spdNFSe.Impressao_CriarDatasets (aXML)
'    Else
'        NomeArquivo = ReturnNumbersOnly(TBFI!IM) & "_" & txtRPS
'        Set ArqTXT = GerArqPastas.OpenTextFile(Localrel & "\NFSe\Impressao\" & NomeArquivo & ".txt")
'        Familiatext = ArqTXT.ReadLine
'
'        spdNFSe.Impressao_CriarDataSetsLog aXML, Familiatext
'    End If
'    spdNFSe.Impressao_Configurar "BrasaoMunicipio", Localrel & "\NFSe\" & "Arquivos\Templates\Impressao\SaoPaulo\Brasao.jpg"
'    If IsNull(TBFI!Logotipo) = False Then spdNFSe.Impressao_Configurar "LogotipoEmitente", IIf(IsNull(TBFI!Logotipo), "", TBFI!Logotipo)
'    spdNFSe.Impressao_Configurar "Titulo", "PREFEITURA MUNICIPAL DE " & UCase(CidadeNFSe)
'    spdNFSe.Impressao_Configurar "SecretariaResponsavel", "SECRETARIA MUNICIPAL DE FINANÇAS DE " & UCase(CidadeNFSe)
'    spdNFSe.Impressao_Configurar "SubtituloNFSe", "NOTA FISCAL DE SERVIÇOS ELETRÔNICA - NFSe"
'    spdNFSe.Impressao_Configurar "SubtituloRPS", "RECIBO PROVISÓRIO DE SERVIÇO - RPS"
'    spdNFSe.Impressao_Configurar "ArquivoMunicipios", Localrel & "\NFSe\" & "Arquivos\Templates\Impressao\Municipios.txt"
'
'    spdProxyNFSe.ComponenteNFSe.Impressao_Editar
'    spdNFSe.Impressao_SetCampo "CPFCNPJPrestador", ReturnNumbersOnly(TBFI!CNPJ)
'
'    If CidadeNFSe = "Salto" Then
'
'        spdNFSe.Impressao_SetCampo "DataEmissao", IIf(IsNull(TBFI!dt_DataEmissao), "", TBFI!dt_DataEmissao)
'        spdNFSe.Impressao_SetCampo "HoraEmissao", IIf(IsNull(TBFI!Hora_emissao), "", Format(Left(TBFI!Hora_emissao, 8), "H:mm:ss"))
'        spdNFSe.Impressao_SetCampo "RazaoSocialPrestador", IIf(IsNull(TBFI!Razao), "", TBFI!Razao)
'        spdNFSe.Impressao_SetCampo "EnderecoPrestador", IIf(IsNull(TBFI!Tipo_endereco), "", TBFI!Tipo_endereco) & " " & IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco) & ", Nº" & IIf(IsNull(TBFI!Numero), "", TBFI!Numero)
'        spdNFSe.Impressao_SetCampo "BairroPrestador", IIf(IsNull(TBFI!Tipo_bairro), "", TBFI!Tipo_bairro) & " " & IIf(IsNull(TBFI!Bairro), "", TBFI!Bairro)
'        spdNFSe.Impressao_SetCampo "CEPPrestador", IIf(IsNull(TBFI!CEP), "", TBFI!CEP)
'        spdNFSe.Impressao_SetCampo "PaisPrestador", "Brasil"
'        spdNFSe.Impressao_SetCampo "ComplementoPrestador", IIf(IsNull(TBFI!complemento), "", TBFI!complemento)
'        spdNFSe.Impressao_SetCampo "MunicipioPrestador", IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
'        spdNFSe.Impressao_SetCampo "InscricaoEstadualPrestador", IIf(IsNull(TBFI!IE), "", ReturnNumbersOnly(TBFI!IE))
'        spdNFSe.Impressao_SetCampo "PaisTomador", "Brasil"
'
'        Set TBProduto = CreateObject("adodb.recordset")
'        TBProduto.Open "Select Sum(vlriss) as valorISS, ISS from tbl_Detalhes_Nota where ID_Nota = " & txtID_nota & " group by ISS", Conexao, adOpenKeyset, adLockReadOnly
'        If TBProduto.EOF = False Then
'            spdNFSe.Impressao_SetCampo "AliquotaSimplesNacional", IIf(IsNull(TBProduto!ISS), 0, TBProduto!ISS)
'            spdNFSe.Impressao_SetCampo "valorISS", IIf(IsNull(TBProduto!valorISS), 0, TBProduto!valorISS)
'        End If
'        TBProduto.Close
'
'        Set TBTotaisnota = CreateObject("adodb.recordset")
'        TBTotaisnota.Open "Select Total_CSLL_serv, Total_IRRF_serv, Total_INSS_serv, Total_Cofins_serv, Total_PIS_serv, dbl_Valor_Total_Nota_Serv from tbl_Totais_Nota where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
'        If TBTotaisnota.EOF = False Then
'            valor = 0
'            Set TBProduto = CreateObject("adodb.recordset")
'            TBProduto.Open "Select Sum(vlriss) as valorISS from tbl_Detalhes_Nota where ID_Nota = " & txtID_nota & " AND Retencao_ISSQN = 'true'", Conexao, adOpenKeyset, adLockReadOnly
'            If TBProduto.EOF = False Then
'                valor = IIf(IsNull(TBProduto!valorISS), 0, TBProduto!valorISS)
'            End If
'            TBProduto.Close
'
'            reter_PIS = False
'            reter_Cofins = False
'            reter_CSLL = False
'            reter_INSS = False
'            reter_IR = False
'            Set TBProduto = CreateObject("adodb.recordset")
'            TBProduto.Open "Select Retencao_ISSQN, Retencao_PIS, Retencao_Cofins, Retencao_CSLL, Retencao_INSS, Retencao_IRRF from tbl_Detalhes_Nota where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
'            Do While TBProduto.EOF = False
'                If TBProduto!Retencao_PIS = True Then reter_PIS = True
'                If TBProduto!Retencao_Cofins = True Then reter_Cofins = True
'                If TBProduto!Retencao_CSLL = True Then reter_CSLL = True
'                If TBProduto!Retencao_ISSQN = True Then reter_INSS = True
'                If TBProduto!Retencao_IRRF = True Then reter_IR = True
'                TBProduto.MoveNext
'            Loop
'            TBProduto.Close
'            spdNFSe.Impressao_SetCampo "BaseCalculoISS", Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "0.00")
'            spdNFSe.Impressao_SetCampo "ValorNf", Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "0.00")
'            spdNFSe.Impressao_SetCampo "ValorLiquidoNota", Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv - IIf(reter_PIS, TBTotaisnota!Total_PIS_serv, 0) - IIf(reter_Cofins, TBTotaisnota!Total_Cofins_serv, 0) - IIf(reter_INSS, TBTotaisnota!Total_INSS_serv, 0) - IIf(reter_IR, TBTotaisnota!Total_IRRF_serv, 0) - IIf(reter_CSLL, TBTotaisnota!Total_CSLL_serv, 0) - valor, "0.00")
'        End If
'        TBTotaisnota.Close
'    Else
'        spdNFSe.Impressao_SetCampo "CidadePrestadorDescricao", IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
'        spdNFSe.Impressao_SetCampo "EmailPrestador", IIf(IsNull(TBFI!Email), "", TBFI!Email)
'        spdNFSe.Impressao_SetCampo "TelefonePrestador", IIf(IsNull(TBFI!telefone), "", TBFI!telefone)
'        spdNFSe.Impressao_SetCampo "EnderecoPrestador", IIf(IsNull(TBFI!Tipo_endereco), "", TBFI!Tipo_endereco) & " " & IIf(IsNull(TBFI!Endereco), "", TBFI!Endereco) & ", Nº" & IIf(IsNull(TBFI!Numero), "", TBFI!Numero) & " - " & IIf(IsNull(TBFI!Tipo_bairro), "", TBFI!Tipo_bairro) & " " & IIf(IsNull(TBFI!Bairro), "", TBFI!Bairro) & " - CEP:" & IIf(IsNull(TBFI!CEP), "", TBFI!CEP)
'    End If
'    spdNFSe.Impressao_SetCampo "UfPrestador", IIf(IsNull(TBFI!UF), "", TBFI!UF)
'    spdProxyNFSe.ComponenteNFSe.Impressao_Salvar
'End If
'TBFI.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
