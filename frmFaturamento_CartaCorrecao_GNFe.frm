VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_CartaCorrecao_GNFe 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Faturamento - Carta de correção"
   ClientHeight    =   10035
   ClientLeft      =   75
   ClientTop       =   435
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximizado
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   10080
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
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
   Begin VB.Timer Timer_status_CCe 
      Interval        =   10000
      Left            =   7200
      Top             =   6780
   End
   Begin VB.TextBox Txt_ID_nota 
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
      Left            =   8340
      Locked          =   -1  'True
      MouseIcon       =   "frmFaturamento_CartaCorrecao_GNFe.frx":0000
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "ID da NF"
      Top             =   6870
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame Frame5 
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
      Height          =   1395
      Left            =   10140
      TabIndex        =   36
      Top             =   990
      Width           =   5145
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Centralizar
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
         MaxLength       =   60
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Status NFe."
         Top             =   390
         Width           =   4785
      End
      Begin VB.TextBox Txt_chave_acesso 
         Alignment       =   2  'Centralizar
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
         MaxLength       =   60
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Chave de acesso NFe."
         Top             =   950
         Width           =   4785
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2340
         TabIndex        =   38
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Chave de acesso"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1957
         TabIndex        =   37
         Top             =   750
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   30
      Top             =   9120
      Width           =   15225
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Centralizar
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
         Left            =   2730
         TabIndex        =   13
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
         TabIndex        =   14
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2014.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   18
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_GNFe.frx":030A
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
      Begin DrawSuite2014.USButton cmdPagAnt 
         Height          =   315
         Left            =   11220
         TabIndex        =   17
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_GNFe.frx":3AB1
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
      Begin DrawSuite2014.USButton cmdPagIr 
         Height          =   315
         Left            =   10110
         TabIndex        =   15
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
      Begin DrawSuite2014.USButton cmdPagPrim 
         Height          =   315
         Left            =   10680
         TabIndex        =   16
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_GNFe.frx":75BF
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
      Begin DrawSuite2014.USButton cmdPagUlt 
         Height          =   315
         Left            =   12300
         TabIndex        =   19
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_CartaCorrecao_GNFe.frx":B6B3
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   41
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   20
      Top             =   990
      Width           =   10065
      Begin VB.TextBox Txt_destinatario 
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
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Nome do destinatário."
         Top             =   950
         Width           =   6555
      End
      Begin VB.TextBox Txt_serie 
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
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Série."
         Top             =   950
         Width           =   645
      End
      Begin VB.CommandButton Cmd_localizar_NF 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1350
         Picture         =   "frmFaturamento_CartaCorrecao_GNFe.frx":EF40
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Localizar nota fiscal."
         Top             =   950
         Width           =   315
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
         ItemData        =   "frmFaturamento_CartaCorrecao_GNFe.frx":F042
         Left            =   180
         List            =   "frmFaturamento_CartaCorrecao_GNFe.frx":F044
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   4020
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   390
         Width           =   3435
      End
      Begin VB.TextBox txtNF 
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
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Número da nota fiscal."
         Top             =   950
         Width           =   1155
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox txtiddestinatario 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "ID do destinatário."
         Top             =   950
         Width           =   885
      End
      Begin MSComCtl2.DTPicker txtdataemissao 
         Height          =   315
         Left            =   5190
         TabIndex        =   2
         ToolTipText     =   "Data de emissão."
         Top             =   390
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   490536961
         CurrentDate     =   39057
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   1823
         TabIndex        =   35
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2790
         TabIndex        =   28
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7710
         TabIndex        =   27
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "NF"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   26
         Top             =   750
         Width           =   195
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Destinatário"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6172
         TabIndex        =   24
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Código"
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
         Left            =   4402
         TabIndex        =   23
         Top             =   180
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Dt. emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5392
         TabIndex        =   22
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Série"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1912
         TabIndex        =   21
         Top             =   750
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correções"
      Enabled         =   0   'False
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
      Height          =   2955
      Left            =   60
      TabIndex        =   25
      Top             =   2400
      Width           =   15225
      Begin VB.ComboBox cmbCertificado 
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
         ItemData        =   "frmFaturamento_CartaCorrecao_GNFe.frx":F046
         Left            =   180
         List            =   "frmFaturamento_CartaCorrecao_GNFe.frx":F048
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         ToolTipText     =   "UF."
         Top             =   2490
         Width           =   9210
      End
      Begin VB.CheckBox Chk_desconsiderar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desconsiderar valor da nota no valor total faturado dos ultimos 12 meses"
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
         Height          =   210
         Left            =   9465
         TabIndex        =   11
         Top             =   2550
         Width           =   5715
      End
      Begin VB.TextBox txtobs 
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
         Height          =   1935
         Left            =   180
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         ToolTipText     =   "Correções."
         Top             =   270
         Width           =   14880
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Certificado digital"
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
         Index           =   12
         Left            =   4043
         TabIndex        =   43
         Top             =   2280
         Width           =   1485
      End
   End
   Begin DrawSuite2014.USImageList USImageList1 
      Left            =   13770
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_CartaCorrecao_GNFe.frx":F04A
      Count           =   1
   End
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   915
      Left            =   60
      TabIndex        =   29
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1614
      ButtonCount     =   13
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   42
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
      ButtonLeft3     =   84
      ButtonTop3      =   2
      ButtonWidth3    =   44
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Excluir"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Excluir (F4)"
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
      ButtonLeft4     =   130
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
      ButtonLeft5     =   177
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
      ButtonLeft6     =   239
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
      ButtonLeft7     =   296
      ButtonTop7      =   2
      ButtonWidth7    =   55
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Enviar"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Enviar CCe (F7)"
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
      ButtonLeft8     =   353
      ButtonTop8      =   2
      ButtonWidth8    =   43
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "E-mail"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Reenciar e-mail."
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
      ButtonLeft9     =   398
      ButtonTop9      =   2
      ButtonWidth9    =   36
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonAlignment10=   2
      ButtonType10    =   1
      ButtonStyle10   =   -1
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState10   =   -1
      ButtonLeft10    =   436
      ButtonTop10     =   4
      ButtonWidth10   =   2
      ButtonHeight10  =   50
      ButtonCaption11 =   "Ajuda"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Ajuda (F1)"
      ButtonKey11     =   "11"
      ButtonAlignment11=   2
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft11    =   440
      ButtonTop11     =   2
      ButtonWidth11   =   41
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonCaption12 =   "Sair"
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonToolTipText12=   "Sair (Esc)"
      ButtonKey12     =   "12"
      ButtonAlignment12=   2
      BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft12    =   483
      ButtonTop12     =   2
      ButtonWidth12   =   30
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
      ButtonEnabled13 =   0   'False
      ButtonKey13     =   "13"
      BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState13   =   5
      ButtonLeft13    =   515
      ButtonTop13     =   2
      ButtonWidth13   =   24
      ButtonHeight13  =   24
      ButtonUseMaskColor13=   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3735
      Left            =   60
      TabIndex        =   12
      Top             =   5370
      Width           =   15225
      _ExtentX        =   26855
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
         Text            =   "Código"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Empresa"
         Object.Width           =   3792
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
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Série"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   7629
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   5115
      EndProperty
   End
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   34
      Top             =   9750
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmFaturamento_CartaCorrecao_GNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Carta As Boolean 'OK
Public StrSql_Localizar_Carta As String 'OK
Dim TBLISTA_Carta As ADODB.Recordset 'OK
Dim emailNF As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=A8dBCFhbghI&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=2&feature=plcp")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBLISTA!ID_empresa) = False And TBLISTA!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBLISTA!ID_empresa
txtCodigo = TBLISTA!ID
txtdataemissao.Value = IIf(IsNull(TBLISTA!Data_emissao), Date, TBLISTA!Data_emissao)
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
Txt_ID_nota = TBLISTA!ID_nota
txtNF = TBLISTA!int_NotaFiscal
Txt_serie = IIf(IsNull(TBLISTA!Serie), "", TBLISTA!Serie)
txtiddestinatario = IIf(IsNull(TBLISTA!Id_Int_Cliente), "", TBLISTA!Id_Int_Cliente)
Txt_destinatario = IIf(IsNull(TBLISTA!txt_Razao_Nome), "", TBLISTA!txt_Razao_Nome)
txtStatus = IIf(IsNull(TBLISTA!status), "Não enviado", TBLISTA!status)

If IsNull(TBLISTA!Chave_acesso) = True Then
    Set TBNivel1 = CreateObject("adodb.recordset")
    TBNivel1.Open "Select Chave_acesso from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & IIf(Txt_ID_nota = "", 0, Txt_ID_nota), Conexao, adOpenKeyset, adLockReadOnly
    If TBNivel1.EOF = False Then Txt_chave_acesso = IIf(IsNull(TBNivel1!Chave_acesso), "", TBNivel1!Chave_acesso)
    TBNivel1.Close
Else
    Txt_chave_acesso = IIf(IsNull(TBLISTA!Chave_acesso), "", TBLISTA!Chave_acesso)
End If

txtobs = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
If TBLISTA!Desconsiderar_valor = True Then Chk_desconsiderar.Value = 1 Else Chk_desconsiderar.Value = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtCodigo.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where order by CC.id", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtCodigo.Text)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimparCampos
        ProcCarregaDados
    Else
        MsgBox ("Fim dos cadastros de carta de correção."), vbInformation
    End If
End If
Novo_Carta = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Cmd_localizar_NF_Click()
On Error GoTo tratar_erro

Minuta = False
Faturamento = False
frmMinuta_notas.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtCodigo.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where order by CC.id", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.BOF = False Then
    TBLISTA.Find (" id =" & txtCodigo.Text)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimparCampos
        ProcCarregaDados
    Else
        MsgBox ("Fim dos cadastros de carta de correção."), vbInformation
    End If
End If
Novo = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Carta.AbsolutePage <> 2 Then
    If TBLISTA_Carta.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Carta.PageCount - 1)
    Else
        TBLISTA_Carta.AbsolutePage = TBLISTA_Carta.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Carta.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = DS_RetornarNumeros(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Carta.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Carta.AbsolutePage)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Carta.AbsolutePage = 1
ProcExibePagina (TBLISTA_Carta.AbsolutePage)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Carta.AbsolutePage <> -3 Then
    If TBLISTA_Carta.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Carta.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Carta.PageCount)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If DS_RetornarNumeros(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Carta.AbsolutePage = TBLISTA_Carta.PageCount
ProcExibePagina (TBLISTA_Carta.AbsolutePage)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: procEnviar
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 13, True
Formulario = "Faturamento/Carta de correção"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
Timer_status_CCe.Enabled = True

' Instanciando o componente NFE
Set spdNFe = New NFeX.spdNFeX
Set spdNFeDataSet = New NFeDataSetX.spdNFeDataSetX

ProcRemoveObjetosResize Me
procCarregaCertificado cmbCertificado

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtCodigo = "" Then
    MsgBox ("Informe a carta de correção antes de visualizar impressão."), vbExclamation
    Exit Sub
End If
If cmbCertificado = "" Then
    MsgBox ("Informe o certificado antes de visualizar impressão."), vbExclamation
    Exit Sub
End If
If Txt_chave_acesso = "" Then
    MsgBox ("Nota fiscal sem chave de acesso."), vbExclamation
    Exit Sub
End If
ProcConfigNFeTECNOSPEED Txt_ID_nota, cmbCertificado
ChaveAcesso = Txt_chave_acesso
xml = spdNFe.ConsultarNF(ChaveAcesso)
spdNFe.VisualizarCCe xml

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente excluir esta(s) carta(s) de correção?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from NF_Carta_Correcao WHERE id = " & .ListItems(InitFor)
            'Conexao_NFe.Execute "DELETE from CBD001EVENTOS WHERE CbdEmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and CbdNtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and CbdNtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and CbdEveDescCC = 'Carta de Correcao'"
            'Conexao_NFe.Execute "DELETE from NFE012 WHERE CbdEmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and CbdNtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and CbdNtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and CbdAcao = 'V'"
            'Conexao_NFe.Execute "DELETE from NFE012EVENTOS WHERE CbdEmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and CbdNtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and CbdNtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and CbdAcao = 'V'"
            'Conexao_NFe.Execute "DELETE from NFE001EVENTOS WHERE EmpCodigo = " & .ListItems(InitFor).ListSubItems(7) & " and NtfNumero = " & .ListItems(InitFor).ListSubItems(4) & " and NtfSerie = " & .ListItems(InitFor).ListSubItems(5) & " and EveDescCC = 'Carta de Correcao'"
            
            '==================================
            Modulo = "Faturamento/Carta de correção"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nota fiscal: " & .ListItems(InitFor).ListSubItems(4) & " - Série: " & .ListItems(InitFor).ListSubItems(5) & " - Destinatário: " & .ListItems(InitFor).ListSubItems(6)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe a(s) carta(s) de correção antes de excluir."), vbExclamation
Else
    MsgBox ("Carta(s) de correção excluída(s) com sucesso."), vbInformation
    ProcLimparCampos
    ProcCarregaLista (1)
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Novo_Carta = False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcCancelar()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente cancelar a liberação desta(s) carta(s) de correção?", vbQuestion + vbYesNo) = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where CC.ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Faturamento/Carta de correção"
                Evento = "Cancelar liberação"
                ID_documento = .ListItems(InitFor)
                Documento = "Nota fiscal: " & TBFI!int_NotaFiscal & " - Série: " & TBFI!Serie & " - Destinatário: " & TBFI!txt_Razao_Nome
                Documento1 = ""
                ProcGravaEvento
                '==================================
                
                OF = TBFI!int_NotaFiscal
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from Empresa where Codigo = " & TBFI!ID_empresa & " and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    caminho = TBMaquinas!Caminho_Nfe & "\Empresa " & TBFI!ID_empresa & " - Serie " & TBFI!Serie & " - Nota " & OF & " - Status CCE.bat"
                    Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                    If GerArqPastas.FileExists(caminho) = True Then Kill caminho
                End If
                TBMaquinas.Close
                
                Conexao_NFe.Execute "DELETE from CBD001EVENTOS WHERE CbdEmpCodigo = " & TBFI!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = " & TBFI!Serie & " and CbdEveDescCC = 'Carta de Correcao'"
                Conexao_NFe.Execute "DELETE from NFE012 WHERE CbdEmpCodigo = " & TBFI!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = " & TBFI!Serie & " and CbdAcao = 'V'"
                Conexao_NFe.Execute "DELETE from NFE012EVENTOS WHERE CbdEmpCodigo = " & TBFI!ID_empresa & " and CbdNtfNumero = " & OF & " and CbdNtfSerie = " & TBFI!Serie & " and CbdAcao = 'V'"
                Conexao_NFe.Execute "DELETE from NFE001EVENTOS WHERE EmpCodigo = " & TBFI!ID_empresa & " and NtfNumero = " & OF & " and NtfSerie = " & TBFI!Serie & " and EveDescCC = 'Carta de Correcao'"
                
                TBFI!CODIGO = Null
                TBFI!status = Null
                TBFI!Chave_acesso = Null
                TBFI.Update
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe a(s) carta(s) de correção antes de cancelar a liberação."), vbExclamation
Else
    MsgBox ("Liberação da(s) carta(s) de correção canceladas(s) com sucesso."), vbInformation
    txtStatus = ""
    Txt_chave_acesso = ""
    ProcCarregaLista (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Novo = False
frmFaturamento_CartaCorrecao_Migrate_Localizar.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
ProcLimparCampos
Frame1.Enabled = True
Frame4.Enabled = True
Novo_Carta = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

txtCodigo = ""
txtdataemissao.Value = Date
txtResponsavel = pubUsuario
txtNF = ""
Txt_serie = ""
txtiddestinatario = ""
Txt_destinatario = ""
txtStatus = ""
Txt_chave_acesso = ""
txtobs = ""
Chk_desconsiderar.Value = 0
CodigoLista = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Carta = True Then
    If MsgBox("A carta de correção ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo + vbQuestion) = vbYes Then
        ProcSalvar
        If Novo_Carta = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Carta = False
Timer_status_CCe.Enabled = False
Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtNF = "" Then
    NomeCampo = "o número da nota fiscal"
    ProcVerificaAcao
    txtNF.SetFocus
    Exit Sub
End If
If Txt_serie = "" Then
    NomeCampo = "a série"
    ProcVerificaAcao
    Txt_serie.SetFocus
    Exit Sub
End If
If txtiddestinatario = "" Then
    NomeCampo = "o destinatário"
    Txt_destinatario.SetFocus
    ProcVerificaAcao
    Exit Sub
End If
If txtobs = "" Then
    NomeCampo = "as correções"
    ProcVerificaAcao
    txtobs.SetFocus
    Exit Sub
End If
If Len(txtobs) < 15 Then
    MsgBox ("É necessário informar no mínimo 15 caracteres no campo de correção, favor revisar."), vbExclamation
    txtobs.SetFocus
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from NF_Carta_Correcao where id = " & IIf(txtCodigo = "", 0, txtCodigo), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!ID_nota = Txt_ID_nota
TBGravar!Data_emissao = txtdataemissao
TBGravar!Responsavel = txtResponsavel
If Chk_desconsiderar.Value = 1 Then TBGravar!Desconsiderar_valor = True Else TBGravar!Desconsiderar_valor = False
TBGravar!Obs = txtobs
TBGravar.Update
txtCodigo = TBGravar!ID
TBGravar.Close

If Novo_Carta = True Then
    MsgBox ("Nova carta de correção cadastrada com sucesso."), vbInformation
    Evento = "Novo"
    StrSql_Localizar_Carta = "Select E.Empresa, CC.*, NF.int_NotaFiscal, NF.Serie, NF.txt_Razao_Nome from (NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID) INNER JOIN Empresa E ON E.Codigo = CC.ID_empresa where CC.id = " & txtCodigo
    ProcCarregaLista (1)
Else
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar"
    ProcCarregaLista (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Faturamento/Carta de correção"
ID_documento = txtCodigo
Documento = "Nota fiscal: " & txtNF & " - Série: " & Txt_serie & " - Destinatário: " & Txt_destinatario
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Carta = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcAtualizarStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
data = Date - 120
If MsgBox("Deseja realmente atualizar o status da(s) carta(s) de correção com data de emissão a patir de " & Format(data, "dd/mm/yy") & "?", vbQuestion + vbYesNo) = vbYes Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where CC.Data_emissao >= '" & Format(data, "Short Date") & "' order by CC.ID", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBGravar.RecordCount
        PBLista.Value = 1
        Contador = 0
        TBGravar.MoveFirst
        Do While TBGravar.EOF = False
            OF = TBGravar!int_NotaFiscal
            
            'Verifica cartas de correção geradas para essa nota
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from NF_Carta_Correcao where ID_nota = " & TBGravar!ID_nota & " and ID < " & TBGravar!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Contador2 = TBAbrir.RecordCount + 1
            Else
                Contador2 = 1
            End If
            TBAbrir.Close
            
            Set TBGravar_NFe = CreateObject("adodb.recordset")
            TBGravar_NFe.Open "Select * from NFE012EVENTOS where CbdNtfNumero = " & OF & " and CbdNtfSerie = '" & TBGravar!Serie & "' and CbdAcao = 'V' and CbdEvenSeq = " & Contador2 & " and CbdEveStsRetCod is not null order by CbdNtfNumero, CbdNtfSerie", Conexao_NFe, adOpenKeyset, adLockOptimistic
            If TBGravar_NFe.EOF = False Then
                TBGravar!CODIGO = TBGravar_NFe!CbdEveStsRetCod
                TBGravar!status = TBGravar_NFe!CbdEveStsRetNome
                TBGravar!Chave_acesso = IIf(IsNull(TBGravar_NFe!CbdEveId), "", TBGravar_NFe!CbdEveId)
                TBGravar.Update
                
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from Empresa where Codigo = " & TBGravar!ID_empresa & " and GNFe = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    caminho = TBMaquinas!Caminho_Nfe & "\Empresa " & TBGravar!ID_empresa & " - Serie " & TBGravar!Serie & " - Nota " & OF & " - Status CCE.bat"
                    Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
                    If GerArqPastas.FileExists(caminho) = True Then Kill caminho
                End If
                TBMaquinas.Close
            End If
            TBGravar_NFe.Close
            TBGravar.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBGravar.Close
    MsgBox ("Status das(s) carta(s) de correção atualizado(s) com sucesso."), vbInformation
    '==================================
    Modulo = "Faturamento/Carta de correção"
    Evento = "Atualizar status"
    ID_documento = 0
    Documento = ""
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcCarregaLista (IIf(DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, DS_RetornarNumeros(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Localizar_Carta = "" Then Exit Sub
Set TBLISTA_Carta = CreateObject("adodb.recordset")
TBLISTA_Carta.Open StrSql_Localizar_Carta, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Carta.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Carta.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Carta.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Carta.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Carta.RecordCount - IIf(Pagina > 1, (TBLISTA_Carta.PageSize * (Pagina - 1)), 0), TBLISTA_Carta.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Carta.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Carta!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Carta!Empresa), "", TBLISTA_Carta!Empresa)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Carta!Data_emissao), "", Format(TBLISTA_Carta!Data_emissao, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Carta!Responsavel), "", TBLISTA_Carta!Responsavel)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Carta!int_NotaFiscal), "", TBLISTA_Carta!int_NotaFiscal)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Carta!Serie), "", TBLISTA_Carta!Serie)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Carta!txt_Razao_Nome), "", TBLISTA_Carta!txt_Razao_Nome)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Carta!status), "Não enviado", TBLISTA_Carta!status)
    End With
    TBLISTA_Carta.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Carta.RecordCount
If TBLISTA_Carta.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Carta.PageCount
ElseIf TBLISTA_Carta.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Carta.PageCount & " de: " & TBLISTA_Carta.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Carta.AbsolutePage - 1 & " de: " & TBLISTA_Carta.PageCount
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Timer_status_CCe.Enabled = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "Código" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from NF_Carta_Correcao where ID = " & .ListItems(InitFor) & " AND CODIGO = 135", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then GoTo Proximo
                TBAbrir.Close
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from NF_Carta_Correcao where ID = " & .ListItems(InitFor) & " AND CODIGO = 135", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                MsgBox ("Não é permitido excluir carta de correção, pois a mesma está com status: " & TBAbrir!status & "."), vbExclamation
                .ListItems.Item(InitFor).Checked = False
            End If
            TBAbrir.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC.*, NF.int_NotaFiscal, NF.Serie, NF.id_int_cliente, NF.txt_Razao_Nome, NF.ID_empresa from NF_Carta_Correcao CC INNER JOIN tbl_Dados_Nota_Fiscal NF ON CC.ID_nota = NF.ID where CC.id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimparCampos
    ProcCarregaDados
    'Puxa certificado salvo no cadastro da empresa
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select CertificadoDigital from Empresa where Codigo = " & TBLISTA!ID_empresa, Conexao, adOpenKeyset, adLockReadOnly
    If TBFI.EOF = False Then
        If IsNull(TBFI!CertificadoDigital) = False Then cmbCertificado = TBFI!CertificadoDigital
    End If
    TBFI.Close
Continua:
    CodigoLista = Lista.SelectedItem.Index
End If
TBLISTA.Close
Frame1.Enabled = True
Frame4.Enabled = True
Novo_Carta = False

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        GoTo Continua
    Else
        MsgBox ("Descrição do erro : " + Error()), vbCritical
    End If
End Sub

Private Sub Timer_status_CCe_Timer()
On Error GoTo tratar_erro

If Timer_status_CCe.Enabled = True Then ProcAtualizaStatusCCe
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: procEnviar
    Case 9: procReenviarEmail
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procEnviar()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
If txtCodigo = "" Then
    MsgBox ("Informe a carta de correção antes de enviar."), vbExclamation
    Exit Sub
End If
If cmbCertificado = "" Then
    MsgBox "Informe o certificado antes de enviar.", vbExclamation
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Status from tbl_dados_nota_fiscal_NFe where id_nota = " & Txt_ID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    If TBAbrir!status <> "100" Or Txt_chave_acesso = "" Then
        TBAbrir.Close
        MsgBox ("Não é permitido enviar, pois a nota fiscal ainda não foi autorizada pelo Sefaz ou esta cancelada."), vbExclamation
        Exit Sub
    End If
End If
TBAbrir.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from NF_Carta_Correcao where ID = " & txtCodigo & " AND CODIGO = 135", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    MsgBox ("Esta carta de correção já foi vinculada a nota fiscal."), vbExclamation
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

ProcConfigNFeTECNOSPEED Txt_ID_nota, cmbCertificado
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select E.UF, E.cidade, NFE.Chave_acesso, CC.OBS from (Empresa E INNER JOIN tbl_Dados_Nota_Fiscal NF ON E.Codigo = NF.ID_empresa) INNER JOIN NF_Carta_Correcao CC ON CC.ID_Nota = NF.ID INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE ON NF.ID = NFE.ID_nota where NF.ID = " & Txt_ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    codUF = FunVerificaCodUF(TBAbrir!Cidade, TBAbrir!UF)
    retorno = spdNFe.EnviarCCe(TBAbrir!Chave_acesso, txtobs, Format(Now(), "yyyy-mm-ddThh:mm:ss"), codUF, "1", 1, FunVerifFusoHorario(False))
    cStat = Mid$(retorno, InStrRev(retorno, "<cStat>") + 7, 3)
        
    If cStat <> "135" Then
        MsgBox "Não foi possível vincular a CCe a NFe, verifique o log para mais detalhes.", vbExclamation
    Else
        Conexao.Execute "Update NF_Carta_Correcao Set Chave_acesso = '" & TBAbrir!Chave_acesso & "', Status = 'Evento registrado e vinculado a NF-e', CODIGO = 135 where ID_nota = " & Txt_ID_nota
        MsgBox "Carta de correção vinculada a NFe com sucesso.", vbInformation
        xml = spdNFe.ConsultarNF(Txt_chave_acesso)
        spdNFe.ImprimirCCe xml, " "
        emailNF = funcSomaEmailNF(Txt_ID_nota)
        ProcEnviarEmailTecnoSpeed emailNF, 3, False, txtNF, Txt_chave_acesso, Txt_serie
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procReenviarEmail()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If

If Txt_ID_nota = 0 Or Txt_ID_nota = "" Then
    MsgBox "Informe a nota fiscal antes de reenviar o e-mail.", vbExclamation
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Status from tbl_dados_nota_fiscal_NFe where id_nota = " & Txt_ID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    If TBAbrir!status <> "100" Then
        TBAbrir.Close
        MsgBox ("Não é permitido reenviar e-mail, pois a nota fiscal ainda não foi autorizada pelo Sefaz ou esta cancelada."), vbExclamation
        Exit Sub
    End If
End If
TBAbrir.Close

If cmbCertificado = "" Then
    MsgBox "Informe o certificado antes de reenviar o e-mail.", vbExclamation
    Exit Sub
End If

ProcConfigNFeTECNOSPEED Txt_ID_nota, cmbCertificado
emailNF = funcSomaEmailNF(Txt_ID_nota)
ProcEnviarEmailTecnoSpeed emailNF, 3, True, txtNF, Txt_chave_acesso, Txt_serie

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub



