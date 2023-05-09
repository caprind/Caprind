VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContas_Receber 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Contas a receber"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   375
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
   Icon            =   "frmContas_Receber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
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
   Begin VB.TextBox txtidintconta 
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
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   2130
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      ToolTipText     =   "Número da conta."
      Top             =   6090
      Visible         =   0   'False
      Width           =   1200
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
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   76
      Top             =   8610
      Width           =   15195
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
         ItemData        =   "frmContas_Receber.frx":1042
         Left            =   6960
         List            =   "frmContas_Receber.frx":105E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   187
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
         Left            =   2730
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
         DibPicture      =   "frmContas_Receber.frx":10D0
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
         DibPicture      =   "frmContas_Receber.frx":4874
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
         DibPicture      =   "frmContas_Receber.frx":837D
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
         DibPicture      =   "frmContas_Receber.frx":C46C
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
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   88
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   5610
         TabIndex        =   86
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   84
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   78
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
         TabIndex        =   77
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   795
      Left            =   55
      TabIndex        =   67
      Top             =   9210
      Width           =   15195
      Begin VB.TextBox txtTotalDevolver 
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
         Height          =   315
         Left            =   10220
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Total a devolver."
         Top             =   390
         Width           =   1550
      End
      Begin VB.TextBox txtTotalAntecipado 
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
         Height          =   315
         Left            =   8660
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Total antecipado."
         Top             =   390
         Width           =   1550
      End
      Begin VB.TextBox Txt_total_descontado 
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
         Height          =   315
         Left            =   11780
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   47
         ToolTipText     =   "Total descontado."
         Top             =   390
         Width           =   1620
      End
      Begin VB.TextBox Txt_total_receber 
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
         Height          =   315
         Left            =   7020
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   44
         ToolTipText     =   "Total a receber."
         Top             =   390
         Width           =   1620
      End
      Begin VB.TextBox txtvalortotal 
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
         Height          =   315
         Left            =   13410
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   48
         ToolTipText     =   "Total geral."
         Top             =   390
         Width           =   1560
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   74
         Top             =   330
         Width           =   6675
         _ExtentX        =   11774
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
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total devolver"
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
         Left            =   10380
         TabIndex        =   87
         Top             =   180
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total antecipado"
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
         Left            =   8715
         TabIndex        =   85
         Top             =   180
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total descontado"
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
         Index           =   0
         Left            =   11850
         TabIndex        =   72
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a receber"
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
         Left            =   7185
         TabIndex        =   71
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral"
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
         Left            =   13763
         TabIndex        =   68
         Top             =   180
         Width           =   2280
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3435
      Left            =   55
      TabIndex        =   53
      Top             =   990
      Width           =   15195
      Begin VB.CommandButton Cmd_localizar_contatos 
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
         Left            =   10260
         Picture         =   "frmContas_Receber.frx":FCF8
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Localizar contatos."
         Top             =   975
         Width           =   315
      End
      Begin VB.CheckBox Chk_antecipacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Antecipação"
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
         Left            =   180
         TabIndex        =   31
         Top             =   2430
         Width           =   1365
      End
      Begin VB.CheckBox Chk_devolucao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Devolução"
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
         Left            =   180
         TabIndex        =   32
         Top             =   2850
         Width           =   1365
      End
      Begin VB.CommandButton Cmd_valor 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1360
         Picture         =   "frmContas_Receber.frx":1000C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Filtrar por cliente."
         Top             =   975
         Width           =   315
      End
      Begin VB.CommandButton cmdtipo 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5895
         Picture         =   "frmContas_Receber.frx":10427
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Filtrar por tipo do documento."
         Top             =   375
         Width           =   315
      End
      Begin VB.ComboBox cmbtipo_conta 
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
         ItemData        =   "frmContas_Receber.frx":10842
         Left            =   4815
         List            =   "frmContas_Receber.frx":10844
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Tipo do documento."
         Top             =   375
         Width           =   1065
      End
      Begin VB.CommandButton Cmd_localizar_tipo_dcto 
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
         Left            =   6225
         Picture         =   "frmContas_Receber.frx":10846
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Localizar tipo do documento."
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton Cmd_data_transacao 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4425
         Picture         =   "frmContas_Receber.frx":10948
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Filtrar por data da transação."
         Top             =   375
         Width           =   315
      End
      Begin VB.ComboBox Cmb_tipo 
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
         ItemData        =   "frmContas_Receber.frx":10D63
         Left            =   1755
         List            =   "frmContas_Receber.frx":10D73
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Tipo."
         Top             =   975
         Width           =   1905
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
         ItemData        =   "frmContas_Receber.frx":10DAF
         Left            =   180
         List            =   "frmContas_Receber.frx":10DB1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   370
         Width           =   3015
      End
      Begin VB.CommandButton CmdForma 
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
         Left            =   7920
         Picture         =   "frmContas_Receber.frx":10DB3
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Localizar forma da baixa."
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton cmdstatus 
         BackColor       =   &H00C0C0C0&
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
         Left            =   14700
         Picture         =   "frmContas_Receber.frx":10EB5
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Filtrar por status."
         Top             =   1560
         Width           =   315
      End
      Begin VB.ComboBox cmb_forma 
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
         ItemData        =   "frmContas_Receber.frx":112D0
         Left            =   4440
         List            =   "frmContas_Receber.frx":1130A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Forma da baixa prevista."
         Top             =   1560
         Width           =   3465
      End
      Begin VB.TextBox txtDocumento 
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
         Left            =   6615
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Número do documento."
         Top             =   370
         Width           =   1470
      End
      Begin VB.ComboBox cmbBanco 
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
         ItemData        =   "frmContas_Receber.frx":11418
         Left            =   180
         List            =   "frmContas_Receber.frx":1141A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Instituição bancária prevista."
         Top             =   1560
         Width           =   4245
      End
      Begin VB.CommandButton Cmdlocalizarcliente 
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
         Left            =   9930
         Picture         =   "frmContas_Receber.frx":1141C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Localizar cliente."
         Top             =   975
         Width           =   315
      End
      Begin VB.TextBox txtIdcliente 
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
         Left            =   3675
         TabIndex        =   19
         ToolTipText     =   "Código do cliente."
         Top             =   975
         Width           =   810
      End
      Begin VB.TextBox txtobservacao 
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
         Height          =   1095
         Left            =   1650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         ToolTipText     =   "Observações."
         Top             =   2190
         Width           =   5820
      End
      Begin VB.ComboBox txtProposta 
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
         Left            =   9585
         TabIndex        =   9
         ToolTipText     =   "Número do pedido interno."
         Top             =   370
         Width           =   1095
      End
      Begin VB.TextBox txtstatus 
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
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   1560
         Width           =   6360
      End
      Begin VB.CommandButton cmdemissao 
         BackColor       =   &H00C0C0C0&
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
         Left            =   12270
         Picture         =   "frmContas_Receber.frx":1151E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Filtrar por data de emissão."
         Top             =   370
         Width           =   315
      End
      Begin VB.CommandButton cmdvencimento 
         BackColor       =   &H00C0C0C0&
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
         Left            =   13875
         Picture         =   "frmContas_Receber.frx":11939
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Filtrar por data de vencimento."
         Top             =   370
         Width           =   315
      End
      Begin VB.CommandButton cmdproposta 
         BackColor       =   &H00C0C0C0&
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
         Left            =   10680
         Picture         =   "frmContas_Receber.frx":11D54
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Filtrar por número do pedido interno."
         Top             =   370
         Width           =   315
      End
      Begin VB.CommandButton cmddoc 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9195
         Picture         =   "frmContas_Receber.frx":1216F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Filtrar por número da nota fiscal."
         Top             =   370
         Width           =   315
      End
      Begin VB.CommandButton cmdLocalizar_fornecedor 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9600
         Picture         =   "frmContas_Receber.frx":1258A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Filtrar por cliente."
         Top             =   975
         Width           =   315
      End
      Begin VB.TextBox txtNFiscal 
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
         Left            =   8100
         MaxLength       =   9
         TabIndex        =   7
         ToolTipText     =   "Número da nota fiscal."
         Top             =   370
         Width           =   1095
      End
      Begin VB.TextBox txtNome_Razao 
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
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "Nome do cliente."
         Top             =   975
         Width           =   5085
      End
      Begin VB.TextBox txtValor 
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
         Height          =   315
         Left            =   180
         TabIndex        =   16
         ToolTipText     =   "Valor."
         Top             =   975
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker mskEmissao 
         Height          =   315
         Left            =   11070
         TabIndex        =   11
         ToolTipText     =   "Data de emissão."
         Top             =   375
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   206176259
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker mskVencimento 
         Height          =   315
         Left            =   12660
         TabIndex        =   13
         ToolTipText     =   "Data de vencimento."
         Top             =   375
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   206176259
         CurrentDate     =   39057
      End
      Begin VB.TextBox txtCidade 
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
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Cidade."
         Top             =   975
         Width           =   3900
      End
      Begin VB.TextBox cbo_UF 
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
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   14595
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "UF."
         Top             =   975
         Width           =   420
      End
      Begin MSMask.MaskEdBox txtParcela 
         Height          =   315
         Left            =   14265
         TabIndex        =   15
         ToolTipText     =   "Número da parcela."
         Top             =   375
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView Lista_PC 
         Height          =   1095
         Left            =   7515
         TabIndex        =   34
         Top             =   2190
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   1931
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   6544
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker Txt_data_transacao 
         Height          =   315
         Left            =   3210
         TabIndex        =   1
         ToolTipText     =   "Data da transação."
         Top             =   370
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   206176259
         CurrentDate     =   39057
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo docto.*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4890
         TabIndex        =   83
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. transação"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3315
         TabIndex        =   82
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contas contábeis"
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
         Left            =   10530
         TabIndex        =   80
         Top             =   1980
         Width           =   1470
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2557
         TabIndex        =   79
         Top             =   780
         Width           =   300
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa*"
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
         Left            =   1267
         TabIndex        =   73
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma da baixa prevista"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5302
         TabIndex        =   69
         Top             =   1350
         Width           =   1740
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6840
         TabIndex        =   66
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Instituição bancária prevista"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1275
         TabIndex        =   65
         Top             =   1350
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4035
         TabIndex        =   64
         Top             =   1980
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor*"
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
         Left            =   495
         TabIndex        =   63
         Top             =   780
         Width           =   540
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. vencto."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   12840
         TabIndex        =   62
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11250
         TabIndex        =   61
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   11228
         TabIndex        =   60
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8265
         TabIndex        =   59
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ped. interno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9690
         TabIndex        =   58
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6750
         TabIndex        =   57
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela*"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   14333
         TabIndex        =   54
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12383
         TabIndex        =   56
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   14708
         TabIndex        =   55
         Top             =   780
         Width           =   195
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar contas"
      Height          =   535
      Left            =   55
      TabIndex        =   70
      Top             =   4410
      Width           =   15195
      Begin VB.ComboBox cmbAno 
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
         Height          =   300
         ItemData        =   "frmContas_Receber.frx":129A5
         Left            =   14250
         List            =   "frmContas_Receber.frx":129A7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         ToolTipText     =   "Ano."
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton OptAteomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
         Height          =   195
         Left            =   1020
         TabIndex        =   50
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton OptDomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
         Height          =   195
         Left            =   150
         TabIndex        =   49
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   345
         Left            =   2160
         TabIndex        =   51
         Top             =   240
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   609
         MultiRow        =   -1  'True
         TabMinWidth     =   1439
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   13
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jan"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fev"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mar"
               Key             =   "Mar"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Abril"
               Key             =   "Abr"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Maio"
               Key             =   "Maio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jun"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jul"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ago"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Set"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Out"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Nov"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dez"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Vencidas"
               Key             =   "Vencidas"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   75
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   20
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
      ButtonCaption6  =   "C. contábil"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Conta contábil (F6)"
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
      ButtonWidth6    =   66
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Agenda"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Agenda do dia (F7)"
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
      ButtonLeft7     =   307
      ButtonTop7      =   2
      ButtonWidth7    =   51
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Parcelar"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Parcelar (F8)"
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
      ButtonLeft8     =   360
      ButtonTop8      =   2
      ButtonWidth8    =   55
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Copiar"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Copiar (F9)"
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
      ButtonLeft9     =   417
      ButtonTop9      =   2
      ButtonWidth9    =   44
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Baixar"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Baixar (F10)"
      ButtonKey10     =   "10"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft10    =   463
      ButtonTop10     =   2
      ButtonWidth10   =   44
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Recomprar"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Recomprar duplicata."
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
      ButtonLeft11    =   509
      ButtonTop11     =   2
      ButtonWidth11   =   71
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonCaption12 =   "Boleto"
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonToolTipText12=   "Emitir boleto (F11)"
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
      ButtonLeft12    =   582
      ButtonTop12     =   2
      ButtonWidth12   =   44
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
      ButtonCaption13 =   "Status"
      ButtonEnabled13 =   0   'False
      ButtonIconSize13=   32
      ButtonToolTipText13=   "Status (F12)"
      ButtonKey13     =   "13"
      ButtonAlignment13=   2
      BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft13    =   628
      ButtonTop13     =   2
      ButtonWidth13   =   45
      ButtonHeight13  =   21
      ButtonUseMaskColor13=   0   'False
      ButtonCaption14 =   "Visualizar"
      ButtonEnabled14 =   0   'False
      ButtonIconSize14=   32
      ButtonToolTipText14=   "Visualizar contas relacionadas."
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
      ButtonLeft14    =   675
      ButtonTop14     =   2
      ButtonWidth14   =   52
      ButtonHeight14  =   21
      ButtonUseMaskColor14=   0   'False
      ButtonCaption15 =   "Retorno"
      ButtonEnabled15 =   0   'False
      ButtonIconSize15=   32
      ButtonToolTipText15=   "Receber arquivo retorno."
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
      ButtonLeft15    =   729
      ButtonTop15     =   2
      ButtonWidth15   =   54
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft16    =   785
      ButtonTop16     =   2
      ButtonWidth16   =   59
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
      ButtonLeft17    =   846
      ButtonTop17     =   4
      ButtonWidth17   =   2
      ButtonHeight17  =   54
      ButtonCaption18 =   "Ajuda"
      ButtonEnabled18 =   0   'False
      ButtonIconSize18=   32
      ButtonToolTipText18=   "Ajuda (F1)"
      ButtonKey18     =   "18"
      ButtonAlignment18=   2
      BeginProperty ButtonFont18 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft18    =   850
      ButtonTop18     =   2
      ButtonWidth18   =   41
      ButtonHeight18  =   21
      ButtonUseMaskColor18=   0   'False
      ButtonCaption19 =   "Sair"
      ButtonEnabled19 =   0   'False
      ButtonIconSize19=   32
      ButtonToolTipText19=   "Sair (Esc)"
      ButtonKey19     =   "19"
      ButtonAlignment19=   2
      BeginProperty ButtonFont19 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft19    =   893
      ButtonTop19     =   2
      ButtonWidth19   =   30
      ButtonHeight19  =   21
      ButtonUseMaskColor19=   0   'False
      ButtonEnabled20 =   0   'False
      ButtonIconSize20=   32
      ButtonKey20     =   "20"
      ButtonAlignment20=   2
      BeginProperty ButtonFont20 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState20   =   5
      ButtonLeft20    =   925
      ButtonTop20     =   2
      ButtonWidth20   =   24
      ButtonHeight20  =   24
      ButtonUseMaskColor20=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   14400
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmContas_Receber.frx":129A9
         Count           =   1
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1410
      Top             =   5700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3650
      Left            =   60
      TabIndex        =   35
      Top             =   4960
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6429
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. venc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   3678
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "N. boleto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Enviado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Remessa"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "IDempresa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "N. duplicata"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmContas_Receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Receber                         As Boolean 'OK
Public StrSql_Contas_Receber                As String 'OK
Public StrSql_Contas_ReceberTotal           As String 'OK
Public StrSql_Contas_Receber_AntecTotal     As String 'OK
Public StrSql_Contas_Receber_DevTotal       As String 'OK
Public StrSql_Contas_ReceberDescTotal       As String 'OK
Public FormulaRel_Contas_Receber            As String 'OK
Dim TBLISTA_Contas_Receber                  As ADODB.Recordset 'OK
Dim ClienteRec                              As String 'OK
Dim IDduplicataRec                          As Long 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=gFKl1cOV_zg&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=30&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaVariaveisCarregaLista()
On Error GoTo tratar_erro

ValorTotal = 0
Valor_total = 0
Valor1 = 0
Codproduto = 0
Dataini = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_antecipacao_Click()
On Error GoTo tratar_erro

If Chk_antecipacao.Value = 1 Then Chk_devolucao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Chk_devolucao_Click()
On Error GoTo tratar_erro

If Chk_devolucao.Value = 1 Then Chk_antecipacao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
ProcCarregaComboBanco

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmitirBoleto()
On Error GoTo tratar_erro

Financeiro_Contas_Receber = True
If Cmb_opcao_lista = "Enviar boleto" Or Cmb_opcao_lista = "Gerar arquivo remessa" Then
    If Cmb_opcao_lista = "Enviar boleto" Then
        MsgTexto = "enviar boleto"
        MsgTexto1 = "Boleto(s) enviado(s)"
        MsgTexto2 = "boleto(s) não foi(ram) enviado(s)."
        Sit_REG = 1
    Else
        MsgTexto = "gerar arquivo remessa"
        MsgTexto1 = "Arquivos(s) remessa gerado(s)"
        MsgTexto2 = "arquivos(s) remessa não foi(ram) gerado(s)."
        Sit_REG = 3
    End If
    Permitido = False
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente " & MsgTexto & " desta(s) contas?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                End If
                Permitido = True
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    ProcCarregaDados
                End If
                frmFaturamento_Prod_serv_boleto.Show
            End If
        Next InitFor
    End With
    If Permitido = False Then
        USMsgBox ("Informe a(s) conta(s) antes de " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
    Else
        If Permitido1 = True Then
            USMsgBox (MsgTexto1 & " com sucesso."), vbInformation, "CAPRIND v5.0"
            ProcCarregaLista (1)
            Lista.SetFocus
            Novo_Receber = False
        Else
            USMsgBox ("O(s) " & MsgTexto2), vbExclamation, "CAPRIND v5.0"
        End If
    End If
Else
    If Novo_Receber = True Then
        USMsgBox ("Salve a conta antes de gerar o boleto"), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Acao = "gerar o boleto"
    If txtidintconta = "" Then
        NomeCampo = "a conta"
        ProcVerificaAcao
        Exit Sub
    End If
    If cmbBanco <> "" Then FamiliaAntiga = cmbBanco.ItemData(cmbBanco.ListIndex) Else FamiliaAntiga = ""
    
    If ProcVerifDadosBoleto(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), IIf(txtidintconta = "", 0, txtidintconta), FamiliaAntiga, 0, "", "emitir", True) = False Then Exit Sub
    
    Sit_REG = 2
    
    frmFaturamento_Prod_serv_boleto.Show
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function ProcVerifDadosBoleto(ID_empresa As Long, ID_conta As Long, ID_banco As String, ID_dest As Long, Tipo_dest As String, Operacao As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifDadosBoleto = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and (Registro_boleto IS NULL or Registro_boleto = N'')", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Operacao & " boleto, pois a empresa não possui registro."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    ProcVerifDadosBoleto = False
    Exit Function
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Tipo, Banco, FormaBaixa from tbl_contas_receber where IdIntConta = " & ID_conta, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Tipo <> "CL" And TBAbrir!Tipo <> "FO" Then
        If MostrarMsg = True Then USMsgBox ("Só é permitido " & Operacao & " boleto para cliente ou fornecedor."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    If Operacao = "emitir" And (IsNull(TBAbrir!Banco) = True Or TBAbrir!Banco = "" Or cmbBanco = "") Or Operacao = "enviar" And (IsNull(TBAbrir!Banco) = True Or TBAbrir!Banco = "") Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                NomeCampo = "o banco"
                ProcVerificaAcao
                cmbBanco.SetFocus
            Else
                USMsgBox ("Não é permitido enviar boleto, pois não existe banco cadastrado nesta conta."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    If Operacao = "emitir" And (IsNull(TBAbrir!FormaBaixa) = True Or TBAbrir!FormaBaixa = "" Or cmb_forma = "") Or Operacao = "enviar" And (IsNull(TBAbrir!FormaBaixa) = True Or TBAbrir!FormaBaixa = "") Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                NomeCampo = "a forma da baixa prevista"
                ProcVerificaAcao
                cmb_forma.SetFocus
            Else
                USMsgBox ("Não é permitido enviar boleto, pois não existe forma da baixa prevista cadastrada nesta conta."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    
    FormaBaixaTexto = TBAbrir!FormaBaixa
    If Left(FormaBaixaTexto, 6) <> "Boleto" And Left(FormaBaixaTexto, 6) <> "BOLETO" Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                USMsgBox ("Não é permitido gerar boleto para essa forma de baixa prevista, favor alterar."), vbExclamation, "CAPRIND v5.0"
                cmb_forma.SetFocus
            Else
                USMsgBox ("Não é permitido enviar boleto, pois a forma da baixa prevista cadastrada nesta conta não permite."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        ProcVerifDadosBoleto = False
        Exit Function
    End If
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select txt_Agencia, txt_conta, Codigo_cedente, Codigo_cedente_registrado from tbl_Instituicoes where ID = " & ID_banco, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!txt_Agencia) = True Or TBAbrir!txt_Agencia = "" Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                NomeCampo = "a agência no cadastro do banco"
                ProcVerificaAcao
            Else
                USMsgBox ("Não é permitido enviar boleto, pois não existe agência cadastrada no banco que está vinculado nesta conta."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    If IsNull(TBAbrir!txt_Conta) = True Or TBAbrir!txt_Conta = "" Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                NomeCampo = "a conta no cadastro do banco"
                ProcVerificaAcao
            Else
                USMsgBox ("Não é permitido enviar boleto, pois não existe conta cadastrada no banco que está vinculado nesta conta."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    If IsNull(TBAbrir!codigo_cedente) = True Or TBAbrir!codigo_cedente = "" Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                NomeCampo = "o código do cedente no cadastro do banco"
                ProcVerificaAcao
            Else
                USMsgBox ("Não é permitido enviar boleto, pois não existe código do cedente cadastrado no banco que está vinculado nesta conta."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    If IsNull(TBAbrir!Codigo_cedente_registrado) = True Or TBAbrir!Codigo_cedente_registrado = "" Then
        If MostrarMsg = True Then
            If Operacao = "emitir" Then
                NomeCampo = "o código do cedente reg. no cadastro do banco"
                ProcVerificaAcao
            Else
                USMsgBox ("Não é permitido enviar boleto, pois não existe código do cedente reg. cadastrado no banco que está vinculado nesta conta."), vbExclamation, "CAPRIND v5.0"
            End If
        End If
        ProcVerificaAcao
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
End If
TBAbrir.Close

If ID_dest <> 0 Then
    If Tipo_dest = "CL" Then
        TabelaFiltro = "Clientes_Contatos"
        CampoFiltro = "IDcliente"
        TipoTexto = "C"
        MsgTexto = "cliente"
    Else
        TabelaFiltro = "Contatos_fornecedor"
        CampoFiltro = "IDfornecedor"
        TipoTexto = "F"
        MsgTexto = "fornecedor"
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IdContato from " & TabelaFiltro & " where " & CampoFiltro & " = " & ID_dest & " and Enviar_boleto = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        If MostrarMsg = True Then USMsgBox ("Não é permitido " & Operacao & " boleto, pois o " & MsgTexto & " não possui nenhum contato configurado para envio do boleto."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select idcobranca from clientes_cobranca where idcliente = " & ID_dest & " and Tipo = '" & TipoTexto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        If MostrarMsg = True Then USMsgBox ("Não é permitido " & Operacao & " boleto, pois o " & MsgTexto & " não possui endereço de cobrança cadastrado."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        ProcVerifDadosBoleto = False
        Exit Function
    End If
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Private Sub ProcRecomprar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "recomprar"
If txtidintconta = "" Then
    NomeCampo = "a conta"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus = "TÍTULO EM ABERTO" Then
    USMsgBox ("Não é permitido recomprar duplicata sem estar descontada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente recomprar esta duplicata?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDduplicata, valor_enviado from troca_titulo_valores where n_conta = " & txtidintconta.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select local_troca from troca_titulo where id = " & TBAbrir!IDDuplicata, Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Permitido = True
            ProcCriarContaPagar
            If Permitido = False Then Exit Sub
            
            If USMsgBox("Deseja criar uma nova conta a receber com estes dados?", vbYesNo, "CAPRIND v5.0") = vbYes Then ProcCriarContaReceber
            
            Conexao.Execute "Update tbl_contas_receber Set Logsit = 'S', Status = 'DUPLICATA DESCONTADA RECOMPRADA', Data_pagamento = '" & Date & "', resprec = '" & pubUsuario & "' where IdIntConta = " & txtidintconta
            Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'True' where IDConta = " & txtidintconta & " and TipoConta = 'R'"
            
            Set TBReceber = CreateObject("adodb.recordset")
            TBReceber.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBContas!local_troca & "' and tbl_contas_receber.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Antecipacao = 'False' and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
            If TBReceber.EOF = False Then
                valor = IIf(IsNull(TBReceber!valor), 0, TBReceber!valor)
                NovoValor = Replace(valor, ",", ".")
                Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & TBContas!local_troca & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
            End If
            TBReceber.Close
        End If
        TBContas.Close
    End If
    TBAbrir.Close
    USMsgBox ("Duplicata recomprada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Financeiro/Contas a receber"
    Evento = "Recomprar duplicata"
    ID_documento = txtidintconta
    Documento = "Documento: " & txtDocumento
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarContaPagar()
On Error GoTo tratar_erro

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select I.* from tbl_Instituicoes I INNER JOIN troca_titulo TT ON I.txt_Descricao = TT.local_troca where TT.id = " & TBAbrir!IDDuplicata, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from tbl_ContasPagar", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!Despesas_NF = False
    TBGravar!Antecipacao = False
    TBGravar!Devolucao = False
    TBGravar!Data_transacao = Date
    TBGravar!Dt_emissao = Date
    TBGravar!Dt_emissao = Date
    TBGravar!dt_Pagamento = mskVencimento.Value
    If IsNull(TBAbrir!valor_enviado) = True Or TBAbrir!valor_enviado = 0 Then TBGravar!dbl_valorpagto = txtValor Else TBGravar!dbl_valorpagto = TBAbrir!valor_enviado
    TBGravar!txt_Parcela = "001/001"
    TBGravar!Tipo = "IN"
    TBGravar!Txt_fornecedor = TBFornecedor!Txt_descricao
    TBGravar!int_codforn = TBFornecedor!ID
    TBGravar!txt_ndocumento = txtNFiscal
    TBGravar!Responsavel = pubUsuario
    TBGravar!status = "TÍTULO EM ABERTO"
    TBGravar!Logsit = "N"
    TBGravar!IdContaReceber = txtidintconta
    TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBGravar.Update
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBGravar!IDFluxo), 0, TBGravar!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = True Then TBFluxo.AddNew
    TBFluxo!IDintconta = TBGravar!IDintconta
    TBFluxo!Operacao = "À Debitar"
    TBFluxo!Data = mskVencimento.Value
    TBFluxo!valor = TBGravar!dbl_valorpagto
    TBFluxo!Descricao = TBFornecedor!Txt_descricao
    TBFluxo!status = "N"
    TBFluxo!int_NotaFiscal = txtNFiscal
    TBFluxo!Bloqueado = False
    TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBFluxo.Update
    Conexao.Execute "Update tbl_contasPagar set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBGravar!IDintconta
    TBFluxo.Close
    
    'Grava valor de recompra no bordero
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select IDduplicata from troca_titulo_valores where n_conta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Set TBFluxo = CreateObject("adodb.recordset")
        TBFluxo.Open "Select Vlrtotalrecompra from troca_titulo where id = " & TBFI!IDDuplicata, Conexao, adOpenKeyset, adLockOptimistic
        If TBFluxo.EOF = False Then
            TBFluxo!Vlrtotalrecompra = IIf(IsNull(TBFluxo!Vlrtotalrecompra), 0, TBFluxo!Vlrtotalrecompra) + TBGravar!dbl_valorpagto
            TBFluxo.Update
        End If
        TBFluxo.Close
    End If
    TBFI.Close
    
    TBGravar.Close
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarContaReceber()
On Error GoTo tratar_erro

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from tbl_contas_receber where IdIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from tbl_contas_receber", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
    TBGravar!IdContaRecomprada = txtidintconta
    TBGravar!Antecipacao = TBMaquinas!Antecipacao
    TBGravar!Devolucao = TBMaquinas!Devolucao
    TBGravar!Tipo_doc = TBMaquinas!Tipo_doc
    TBGravar!txt_ndocumento = TBMaquinas!txt_ndocumento
    TBGravar!NFiscal = TBMaquinas!NFiscal
    TBGravar!Proposta = TBMaquinas!Proposta
    TBGravar!Data_transacao = TBMaquinas!Data_transacao
    TBGravar!emissao = TBMaquinas!emissao
    TBGravar!Tipo = TBMaquinas!Tipo
    TBGravar!Nome_Razao = TBMaquinas!Nome_Razao
    TBGravar!IDCliente = TBMaquinas!IDCliente
    TBGravar!Cidade = TBMaquinas!Cidade
    TBGravar!Estado = TBMaquinas!Estado
    TBGravar!valor = TBMaquinas!valor
    TBGravar!status = "TÍTULO EM ABERTO"
    TBGravar!Responsavel = pubUsuario
    TBGravar!Vencimento = TBMaquinas!Vencimento
    TBGravar!Parcela = TBMaquinas!Parcela
    TBGravar!Observacoes = TBMaquinas!Observacoes
    TBGravar!Banco = TBMaquinas!Banco
    TBGravar!FormaBaixa = TBMaquinas!FormaBaixa
    TBGravar!ID_empresa = TBMaquinas!ID_empresa
    TBGravar.Update
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBGravar!IDFluxo), 0, TBGravar!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = True Then TBFluxo.AddNew
    TBFluxo!IDintconta = TBGravar!IDintconta
    TBFluxo!Operacao = "À Creditar"
    TBFluxo!Data = TBGravar!Vencimento
    TBFluxo!valor = TBGravar!valor
    TBFluxo!Descricao = TBGravar!Nome_Razao
    TBFluxo!status = "N"
    TBFluxo!int_NotaFiscal = TBGravar!NFiscal
    TBFluxo!Documento = TBGravar!txt_ndocumento
    TBFluxo!Bloqueado = False
    TBFluxo!ID_empresa = TBGravar!ID_empresa
    TBFluxo.Update
    Conexao.Execute "Update tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBGravar!IDintconta
    TBFluxo.Close
    
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Familia_financeiro where IDConta = " & txtidintconta & " and TipoConta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Familia_financeiro", Conexao, adOpenKeyset, adLockOptimistic
            TBFI.AddNew
            TBFI!ID_PC = TBFIltro!ID_PC
            TBFI!IDConta = TBGravar!IDintconta
            TBFI!valor = TBFIltro!valor
            TBFI!TipoConta = "R"
            TBFI.Update
            TBFI.Close
            TBFIltro.MoveNext
        Loop
    End If
    TBFIltro.Close

    TBGravar.Close
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) conta(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmContas_receber_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPlanoContas()
On Error GoTo tratar_erro
    
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = True
Financeiro_Contas_Pagas = False
Financeiro_Contas_Recebidas = False
If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de cadastrar a(s) conta(s) contábil."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmFamilia_financeiro.Show

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
    .ButtonToolTipText(12) = "Emitir boleto (F11)"
    If Left(Cmb_opcao_lista, 7) = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(10) = 5
        .ButtonState(13) = 5
    ElseIf Cmb_opcao_lista = "Baixar" Then
            .ButtonState(4) = 5
            .ButtonState(10) = 0
            .ButtonState(13) = 5
        ElseIf Cmb_opcao_lista = "Status" Then
                .ButtonState(4) = 5
                .ButtonState(10) = 5
                .ButtonState(13) = 0
            ElseIf Cmb_opcao_lista = "Enviar boleto" Then
                    .ButtonToolTipText(12) = "Enviar boleto (F11)"
                ElseIf Cmb_opcao_lista = "Gerar arquivo remessa" Then
                        .ButtonToolTipText(12) = "Gerar arquivo remessa (F11)"
                    Else
                        .ButtonState(4) = 5
                        .ButtonState(10) = 5
                        .ButtonState(13) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_Click()
On Error GoTo tratar_erro

txtIDcliente = ""
txtNome_Razao = ""
txtCidade = ""
cbo_UF = ""
If Cmb_tipo = "Cliente" Or Cmb_tipo = "Fornecedor" Then Cmd_localizar_contatos.Enabled = True Else Cmd_localizar_contatos.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_data_transacao_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "Data_transacao = '" & Format(Txt_data_transacao.Value, "Short Date") & "'", "{tbl_Contas_receber.Data_transacao} = Date(" & Year(Txt_data_transacao.Value) & "," & Month(Txt_data_transacao.Value) & "," & Day(Txt_data_transacao.Value) & ")", True, True, False, False, Txt_data_transacao, Txt_data_transacao, "Data_transacao"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcFiltrarContas(TextoFiltro As String, TextoFiltroRel As String, Imprimir As Boolean, DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataInicio As Date, DataFinal As Date, Ordenar As String)
On Error GoTo tratar_erro

NomeRel = "Contas_receber.rpt"
ProcConstruirFiltroPadrao TextoFiltro, TextoFiltroRel, True, True
ProcSalvarDadosRel DataTransacao, DataEmissao, DataVencimento, DataInicio, DataFinal
StrSql_Contas_Receber_AntecTotal = ""
StrSql_Contas_Receber_DevTotal = ""
ProcCarregaLista (1)
Imprimir = Imprimir
'Novo_Receber = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_contatos_Click()
On Error GoTo tratar_erro

If txtNome_Razao <> "" Then
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = True
    Financeiro_Contas_Recebidas = False
    If Cmb_tipo = "Fornecedor" Then
        Compras_Cotacao = False
        Compras_Pedido = False
        frmCompras_Pedido_contatos.Show 1
    Else
        Analise_critica = False
        Vendas_Proposta = False
        Vendas_PI = False
        Telemarketing = False
        Qualidade_PPAP_PSW = False
        frmVendas_propostaII_contato.Show 1
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_tipo_dcto_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = True
Clientes = False
Compras_Fornecedores = False
frmContas_Tipo_Dcto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_valor_Click()
On Error GoTo tratar_erro

If txtValor <> "" Then
    valor = txtValor
    NovoValor = Replace(valor, ",", ".")
    ProcFiltrarContas "Valor = " & NovoValor, "{tbl_Contas_receber.Valor} = " & NovoValor, True, False, False, False, Date, Date, "vencimento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdForma_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = False
Financeiro_Forma_Pgto_Pagar = False
Financeiro_Contas_Receber = True
Financeiro_Forma_Pgto_Receber = False
frmContas_Forma_Pagamento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizarCliente_Click()
On Error GoTo tratar_erro

ProcLocalizarCliente

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarCliente()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False
If Cmb_tipo = "Cliente" Then
    frmVendas_LocalizarCliente.Show 1
ElseIf Cmb_tipo = "Fornecedor" Then
        FrmCompras_localizafornecedor.Show 1
    ElseIf Cmb_tipo = "Funcionário" Then
            frmContas_pagar_localizar_func.Show 1
        Else
            frmContas_pagar_localizar_inst.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtidintconta.Text = ""
Txt_data_transacao.Value = Date
cmbtipo_conta.ListIndex = -1
txtDocumento = ""
txtNFiscal.Text = ""
txtProposta.Clear
mskEmissao.Value = Date
txtIDcliente.Text = ""
Cmb_tipo = "Cliente"
txtNome_Razao.Text = ""
txtCidade.Text = ""
cbo_UF.Text = ""
txtValor.Text = ""
mskVencimento.Value = Date
txtparcela.Text = "___/___"
txtStatus.Text = "TÍTULO EM ABERTO"
cmbBanco.ListIndex = -1
cmb_forma.ListIndex = -1
Chk_antecipacao.Value = 0
Chk_devolucao.Value = 0
txtObservacao.Text = ""
CodigoLista = 0
Lista_PC.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarTodas()
On Error GoTo tratar_erro

NomeRel = "Contas_receber.rpt"
ProcConstruirFiltroPadrao "CR.IDintconta IS NOT NULL", "Not(IsNull({tbl_Contas_receber.IDintconta}))", True, True
ProcSalvarDadosRel False, False, False, Date, Date = False
StrSql_Contas_Receber_AntecTotal = ""
StrSql_Contas_Receber_DevTotal = ""
ProcCarregaLista (1)
Imprimir = True
'Novo_Receber = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_Contas_Receber = CreateObject("adodb.recordset")
TBLISTA_Contas_Receber.Open StrSql_Contas_Receber, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Contas_Receber.EOF = False Then ProcExibePagina (Pagina)
ProcCarregaTotal
'Debug.print StrSql_Contas_Receber

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

IDduplicataRec = 0
ClienteRec = ""

ProcLimpaVariaveisCarregaLista
Lista.ListItems.Clear
TBLISTA_Contas_Receber.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Contas_Receber.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Contas_Receber.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Contas_Receber.RecordCount - IIf(Pagina > 1, (TBLISTA_Contas_Receber.PageSize * (Pagina - 1)), 0), TBLISTA_Contas_Receber.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Contas_Receber.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems.Add(, , TBLISTA_Contas_Receber!IDintconta)
        .SubItems(1) = IIf(IsNull(TBLISTA_Contas_Receber!emissao), Date, Format(TBLISTA_Contas_Receber!emissao, "dd/mm/yy"))
        .SubItems(2) = IIf(IsNull(TBLISTA_Contas_Receber!Vencimento), Date, Format(TBLISTA_Contas_Receber!Vencimento, "dd/mm/yy"))
        
        If TBLISTA_Contas_Receber!Antecipacao = True Then qt = IIf(IsNull(TBLISTA_Contas_Receber!Saldo_antecipacao), 0, TBLISTA_Contas_Receber!Saldo_antecipacao) Else qt = IIf(IsNull(TBLISTA_Contas_Receber!valor), 0, TBLISTA_Contas_Receber!valor)
        .SubItems(3) = Format(qt, "###,##0.00")
        
        .SubItems(4) = IIf(IsNull(TBLISTA_Contas_Receber!txt_ndocumento), "", TBLISTA_Contas_Receber!txt_ndocumento)
        .SubItems(5) = IIf(IsNull(TBLISTA_Contas_Receber!NFiscal), "", TBLISTA_Contas_Receber!NFiscal)
        .SubItems(6) = IIf(IsNull(TBLISTA_Contas_Receber!Parcela), "", TBLISTA_Contas_Receber!Parcela)
        .SubItems(7) = IIf(IsNull(TBLISTA_Contas_Receber!Nome_Razao), "", TBLISTA_Contas_Receber!Nome_Razao)
        .SubItems(8) = IIf(IsNull(TBLISTA_Contas_Receber!Responsavel), "", TBLISTA_Contas_Receber!Responsavel)
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select I.Id as ID_banco, DR.Seq_remessa, DR.Nosso_numero, DR.Data_emissao, DR.Enviado from tbl_Detalhes_Recebimento DR INNER JOIN tbl_Instituicoes I ON DR.txt_Portador_Banco = I.txt_Descricao where DR.IDContaReceber = " & TBLISTA_Contas_Receber!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .SubItems(9) = IIf(IsNull(TBAbrir!Nosso_Numero), "", TBAbrir!Nosso_Numero)
            .SubItems(10) = IIf(TBAbrir!Enviado = True, "Sim", "Não")
            If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" And TBAbrir!ID_banco <> "" Then .SubItems(11) = FunFormataNumeroArqRemessa(TBAbrir!Data_emissao, TBAbrir!ID_banco, TBAbrir!Seq_remessa)
        Else
            .SubItems(9) = ""
            .SubItems(10) = ""
            .SubItems(11) = ""
        End If
        TBAbrir.Close
                        
        .SubItems(12) = IIf(IsNull(TBLISTA_Contas_Receber!ID_empresa), 0, TBLISTA_Contas_Receber!ID_empresa)
        .SubItems(13) = IIf(IsNull(TBLISTA_Contas_Receber!IDDuplicata), "", TBLISTA_Contas_Receber!IDDuplicata)
        Dataini = TBLISTA_Contas_Receber!Vencimento
        If Date > Dataini Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .ListSubItems(3).ForeColor = vbRed
            .ListSubItems(4).ForeColor = vbRed
            .ListSubItems(5).ForeColor = vbRed
            .ListSubItems(6).ForeColor = vbRed
            .ListSubItems(7).ForeColor = vbRed
            .ListSubItems(8).ForeColor = vbRed
            .ListSubItems(9).ForeColor = vbRed
            .ListSubItems(10).ForeColor = vbRed
            .ListSubItems(11).ForeColor = vbRed
            .ListSubItems(12).ForeColor = vbRed
            .ListSubItems(13).ForeColor = vbRed
        End If
    End With
    TBLISTA_Contas_Receber.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Contas_Receber.RecordCount
If TBLISTA_Contas_Receber.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Contas_Receber.PageCount
ElseIf TBLISTA_Contas_Receber.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Receber.PageCount & " de: " & TBLISTA_Contas_Receber.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Receber.AbsolutePage - 1 & " de: " & TBLISTA_Contas_Receber.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

'À receber
valor = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open StrSql_Contas_ReceberTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    valor = IIf(IsNull(TBTotaisnota!TotContas), 0, TBTotaisnota!TotContas)
End If
TBTotaisnota.Close
Txt_total_receber.Text = Format(valor, "###,##0.00")

'Antecipado
Valor1 = 0
If StrSql_Contas_Receber_AntecTotal <> "" Then
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open StrSql_Contas_Receber_AntecTotal, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = False Then
        Valor1 = IIf(IsNull(TBTotaisnota!TotContas1), 0, TBTotaisnota!TotContas1)
    End If
End If
txtTotalAntecipado.Text = Format(Valor1, "###,##0.00")

'Devolver
Valor2 = 0
If StrSql_Contas_Receber_DevTotal <> "" Then
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open StrSql_Contas_Receber_DevTotal, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = False Then
        Valor2 = IIf(IsNull(TBTotaisnota!TotContas), 0, TBTotaisnota!TotContas)
    End If
End If
txtTotalDevolver.Text = Format(Valor2, "###,##0.00")

'Descontado
Valor3 = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open StrSql_Contas_ReceberDescTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    Valor3 = IIf(IsNull(TBTotaisnota!TotContas), 0, TBTotaisnota!TotContas)
End If
Txt_total_descontado = Format(Valor3, "###,##0.00")

'Total geral (A receber - Antecipado + A devolver + Descontado)
qt = valor - Valor1 + Valor2 + Valor3
txtValorTotal.Text = IIf(qt < 0, "0,00", Format(qt, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmddoc_Click()
On Error GoTo tratar_erro

If txtNFiscal.Text <> "" Then
    ProcFiltrarContas "nfiscal = '" & txtNFiscal.Text & "'", "{tbl_Contas_receber.nfiscal} = '" & txtNFiscal.Text & "'", True, False, False, False, Date, Date, "vencimento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdemissao_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "Emissao = '" & Format(mskEmissao.Value, "Short Date") & "'", "{tbl_Contas_receber.Emissao} = Date(" & Year(mskEmissao.Value) & "," & Month(mskEmissao.Value) & "," & Day(mskEmissao.Value) & ")", True, False, True, False, mskEmissao, mskEmissao, "Emissao"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_fornecedor_Click()
On Error GoTo tratar_erro

If txtNome_Razao.Text <> "" Then
    ProcFiltrarContas "nome_razao = '" & txtNome_Razao.Text & "'", "{tbl_Contas_receber.nome_razao} = '" & txtNome_Razao.Text & "'", True, False, False, False, Date, Date, "vencimento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Receber.AbsolutePage <> 2 Then
    If TBLISTA_Contas_Receber.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Contas_Receber.PageCount - 1)
    Else
        TBLISTA_Contas_Receber.AbsolutePage = TBLISTA_Contas_Receber.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Contas_Receber.AbsolutePage)
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
    TBLISTA_Contas_Receber.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Contas_Receber.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Receber.AbsolutePage = 1
ProcExibePagina (TBLISTA_Contas_Receber.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Receber.AbsolutePage <> -3 Then
    If TBLISTA_Contas_Receber.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Contas_Receber.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Contas_Receber.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Receber.AbsolutePage = TBLISTA_Contas_Receber.PageCount
ProcExibePagina (TBLISTA_Contas_Receber.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdproposta_Click()
On Error GoTo tratar_erro

Proposta = True
If txtProposta.Text <> "" Then
    NomeRel = "Contas_receber.rpt"
    ProcConstruirFiltroPadrao "PN.Proposta = '" & txtProposta & "'", "{tbl_proposta_nota.proposta} = '" & txtProposta & "'", True, True
    ProcSalvarDadosRel False, False, False, Date, Date
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open StrSql_Contas_Receber, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        ProcConstruirFiltroPadrao "CR.proposta = '" & txtProposta & "'", "{tbl_contas_receber.proposta} = '" & txtProposta & "'", True, True
    End If
    TBAbrir.Close
    Imprimir = True
Else
    ProcFiltrarTodas
End If
StrSql_Contas_Receber_AntecTotal = ""
StrSql_Contas_Receber_DevTotal = ""
ProcCarregaLista (1)
Proposta = False
Novo_Receber = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro

If txtStatus.Text <> "" Then
    ProcFiltrarContas "Status = '" & txtStatus.Text & "'", "{tbl_Contas_receber.Status} = '" & txtStatus.Text & "'", True, False, False, False, Date, Date, "vencimento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdtipo_Click()
On Error GoTo tratar_erro

If cmbtipo_conta.Text <> "" Then
    ProcFiltrarContas "Tipo_doc = '" & cmbtipo_conta.Text & "'", "{tbl_Contas_receber.Tipo_doc} = '" & cmbtipo_conta.Text & "'", True, False, False, False, Date, Date, "vencimento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdvencimento_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "vencimento = '" & Format(mskVencimento.Value, "Short Date") & "'", "{tbl_Contas_receber.vencimento} = Date(" & Year(mskVencimento.Value) & "," & Month(mskVencimento.Value) & "," & Day(mskVencimento.Value) & ")", True, False, False, True, mskVencimento, mskVencimento, "vencimento"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF6: ProcPlanoContas
    Case vbKeyF7: ProcAgendaDia
    Case vbKeyF8: ProcParcelar
    Case vbKeyF9: ProcCopiar
    Case vbKeyF10: If Cmb_opcao_lista = "Baixar" Then ProcReceber
    Case vbKeyF11: ProcEmitirBoleto
    Case vbKeyF12: If Cmb_opcao_lista = "Status" Then ProcStatus
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 20, True

Formulario = "Financeiro/Contas a receber"
Direitos
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Baixar"
Imprimir = False
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
mskEmissao.Value = Date
mskVencimento.Value = Date
ProcCarregaComboAno cmbAno, Year(Now) - 2, 2
TabFiltro.Tabs(Month(Date)).Selected = True

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboBanco()
On Error GoTo tratar_erro

ProcCarregaComboBancoFinanceiro cmbBanco, "txt_Descricao IS NOT NULL and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloqueado = 'false' and DtValidacao IS NOT NULL", True
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

ProcCarregaComboFormaPgtoRcbto cmb_forma, "Tipo = 'R'"
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaTipoDocumento()
On Error GoTo tratar_erro

ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'R'"
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbBanco.ListIndex = -1
cmb_forma.ListIndex = -1
cmbtipo_conta.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Contas_receber where IdIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Banco) = False And TBAbrir!Banco <> "" Then cmbBanco = TBAbrir!Banco
    If IsNull(TBAbrir!FormaBaixa) = False And TBAbrir!FormaBaixa <> "" Then cmb_forma = TBAbrir!FormaBaixa
    If IsNull(TBAbrir!Tipo_doc) = False And TBAbrir!Tipo_doc <> "" Then cmbtipo_conta = TBAbrir!Tipo_doc
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Contas a receber"
Direitos
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
ProcLimpaVariaveisPrincipais
NomeRel = "Contas_receber.rpt"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Cmb_opcao_lista = "Gerar duplicata" Then
    IDlista = 0
    valor = 0
    Permitido = False
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                Permitido = True
                'Verifica se tem conta selecionada com ID duplicata
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select IDduplicata, Valor from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor) & " and IDduplicata IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    IDlista = TBContas!IDDuplicata
                Else
                    valor = valor + .ListItems(InitFor).ListSubItems(3)
                End If
                TBContas.Close
            End If
        Next InitFor
        
        If Permitido = False Then
            USMsgBox ("Informe a(s) conta(s) antes de gerar duplicata."), vbExclamation, "CAPRIND v5.0"
        Else
            frmContas_receber_duplicatas.Show 1
        End If
    End With
Else
    If Imprimir = True Then
        frmContas_Receber_menuimpressao.Show 1
    Else
        USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362CR" Then frmContas_Receber_atualizar.Show 1

'Rotina para adicionar a CC de juros nas contas recebidas no banco da CAPRIND
'Set TBAbrir = CreateObject("adodb.recordset")
'TBAbrir.Open "SELECT CR.IDIntconta, CR.Nome_Razao, CR.valortitulorecebido, ROUND(SUM(F.Valor), 2) AS ValorCC, ROUND(ROUND(CR.valortitulorecebido, 2) - ROUND(SUM(F.Valor), 2), 2) AS Diferenca FROM tbl_contas_receber AS CR LEFT OUTER JOIN Familia_financeiro AS F ON CR.IDIntconta = F.IDConta WHERE (F.TipoConta = 'R') GROUP BY CR.IDIntconta, CR.valortitulorecebido, CR.Nome_Razao Having (CR.valortitulorecebido > Round(Sum(f.Valor), 2)) AND (ROUND(ROUND(CR.valortitulorecebido, 2) - ROUND(SUM(F.Valor), 2), 2) <= 50) ORDER BY CR.Nome_Razao, CR.IDIntconta", Conexao, adOpenKeyset, adLockReadOnly
'If TBAbrir.EOF = False Then
'    PBLista.Min = 0
'    PBLista.Max = TBAbrir.RecordCount
'    PBLista.Value = 1
'    Contador = 0
'    Do While TBAbrir.EOF = False
'        Set TBGravar = CreateObject("adodb.recordset")
'        TBGravar.Open "SELECT * from Familia_financeiro where IDConta = " & TBAbrir!IDintconta & " and ID_PC = 470 and TipoConta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
'        If TBGravar.EOF = True Then TBGravar.AddNew
'        TBGravar!IDConta = TBAbrir!IDintconta
'        TBGravar!TipoConta = "R"
'        TBGravar!Valor = Format(TBAbrir!Diferenca, "###,##0.00")
'        TBGravar!Pago_recebido = True
'        TBGravar!ID_PC = 470
'        TBGravar.Update
'        TBGravar.Close
'
'        TBAbrir.MoveNext
'        Contador = Contador + 1
'        PBLista.Value = Contador
'    Loop
'End If
'TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmContas_Receber_atualizar
        If .Chk1.Value = 1 Then
            'Status da conta contábil
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber where Status = 'DUPLICATA DESCONTADA RECOMPRADA' order by IdIntConta", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TBAbrir.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBAbrir.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBAbrir.MoveFirst
                Do While TBAbrir.EOF = False
                    Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'True' where IDConta = " & TBAbrir!IDintconta & " and TipoConta = 'R'"
                    TBAbrir.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBAbrir.Close
        End If
        
        If .Chk2.Value = 1 Then
            'ID do cliente
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select idcliente, NomeRazao from clientes order by idcliente", Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                TBClientes.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBClientes.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBClientes.MoveFirst
                Do While TBClientes.EOF = False
                    Conexao.Execute "Update tbl_contas_receber Set IDcliente = " & TBClientes!IDCliente & " where Nome_Razao = '" & TBClientes!NomeRazao & "'"
                    TBClientes.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBClientes.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Dados das contas descontadas
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber where Status = 'DUPLICATA DESCONTADA LIQUIDADA' order by IdIntConta", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TBAbrir.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBAbrir.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBAbrir.MoveFirst
                Do While TBAbrir.EOF = False
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from troca_titulo_valores where n_conta = " & TBAbrir!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = False Then
                        If TBAbrir!Data_pagamento < TBAbrir!Vencimento Then TBAbrir!Data_pagamento = TBAbrir!Vencimento
                        TBAbrir!valortitulorecebido = TBGravar!valor_enviado
                    End If
                    TBGravar.Close
                    TBAbrir.Update
                    TBAbrir.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBAbrir.Close
            Conexao.Execute "Update tbl_contas_receber Set LogSit = 'N', Bloqueado = 'False' where Status = 'DUPLICATA DESCONTADA EM ABERTO'"
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Financeiro/Contas a receber"
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

Private Sub ProcReceber()
On Error GoTo tratar_erro
                
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Contador = 0
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                'Verifica a primeira conta selecionada e informa quais as novas contas poderão ser selecionadas
                If Contador = 0 Then
                    If TBContas!Antecipacao = True Then
                        Contador = 1
                    ElseIf TBContas!Devolucao = True Then
                            Contador = 2
                        ElseIf TBContas!status = "DUPLICATA DESCONTADA EM ABERTO" Then
                                Contador = 3
                            Else
                                Contador = 4
                    End If
                ElseIf Contador = 1 Then
                        If TBContas!Antecipacao = False Then
                            USMsgBox ("Só é permitido baixar conta de antecipação."), vbExclamation, "CAPRIND v5.0"
                            Exit Sub
                        End If
                    ElseIf Contador = 2 Then
                            If TBContas!Devolucao = False Then
                                USMsgBox ("Só é permitido baixar conta de devolução."), vbExclamation, "CAPRIND v5.0"
                                Exit Sub
                            End If
                    ElseIf Contador = 3 Then
                            If TBContas!status <> "DUPLICATA DESCONTADA EM ABERTO" Then
                                USMsgBox ("Só é permitido baixar conta descontada."), vbExclamation, "CAPRIND v5.0"
                                Exit Sub
                            End If
                        Else
                            If TBContas!status = "DUPLICATA DESCONTADA EM ABERTO" Or TBContas!Antecipacao = True Or TBContas!Devolucao = True Then
                                USMsgBox ("Só é permitido baixar conta em aberto ou baixada parcial que não seja conta de antecipação nem devolução."), vbExclamation, "CAPRIND v5.0"
                                Exit Sub
                            End If
                End If
            End If
            TBContas.Close
        End If
    Next InitFor
End With

Permitido1 = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido1 = False Then
                If USMsgBox("Deseja realmente baixar esta(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido1 = True
            GoTo 2
        End If
    Next InitFor
End With
2:
    If Permitido1 = False Then
        USMsgBox ("Informe a(s) conta(s) antes de baixar."), vbExclamation, "CAPRIND v5.0"
    Else
        Permitido1 = False
        frm_Baixas_Receber.Show 1
        If Permitido1 = True Then
            ProcCarregaLista (1)
            ProcLimpaCampos
            Lista.SetFocus
            If Lista.ListItems.Count = 0 Then Exit Sub
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contas_receber where idintconta = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                ProcCarregaDados
                CodigoLista = Lista.SelectedItem.index
            End If
        End If
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAgendaDia()
On Error GoTo tratar_erro

StrSql_Contas_Receber = "Select * from tbl_Contas_RECEBER WHERE vencimento = '" & Format(Date, "Short Date") & "' and logsit='N' and bloqueado = 'False' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Antecipacao = 'False' order by tbl_Contas_receber.vencimento, tbl_Contas_receber.IdIntConta"
NomeRel = "Contas_receber.rpt"
FormulaRel_Contas_Receber = "{tbl_Contas_receber.vencimento} = Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ") and {tbl_Contas_receber.LogSit} ='N' and {tbl_Contas_receber.bloqueado} = False and {tbl_contas_receber.Antecipacao} = False and {tbl_contas_receber.Devolucao} = False and {tbl_contas_receber.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
ProcSalvarDadosRel False, False, False, Date, Date = False
Imprimir = True
StrSql_Contas_Receber_AntecTotal = ""
StrSql_Contas_Receber_DevTotal = ""
ProcCarregaLista (1)
Novo_Receber = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If txtidintconta = "" Or txtValor.Text = "" Or txtIDcliente.Text = "" Then
    USMsgBox ("Informe a conta antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Receber = True Then
    USMsgBox ("Salve a conta antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Chk_antecipacao.Value = 1 Then
    USMsgBox ("Não é permitido copiar esta conta, pois a mesma é uma antecipação."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frm_Contas_parcelamento_receber.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
If Cmb_opcao_lista = "Excluir" Then
    TextoMsg = "conta(s)"
    TextoMsg1 = "Conta(s)"
Else
    TextoMsg = "duplicata(s)"
    TextoMsg1 = "Duplicata(s)"
End If
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) " & TextoMsg & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If Cmb_opcao_lista = "Excluir" Then
                    'Fluxo de Caixa
                    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
                    
                    If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                        'Fluxo de Caixa
                        If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Or TBContas!FormaBaixa = "DOC" Or TBContas!FormaBaixa = "TED" Or TBContas!FormaBaixa = "MALOTE" Or IsNull(TBContas!ID_varias) = False And TBContas!ID_varias > 0 Then
                            TextoFiltroData = "Data = '" & Format(TBContas!Data_movimentacao, "Short Date") & "' and"
                            Select Case TBContas!FormaBaixa
                                Case "CHEQUE":
                                    Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                                    TextoFiltroData = ""
                                Case "CHEQUE PRÉ-DATADO":
                                    Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                                    TextoFiltroData = ""
                                Case "DOC": Cheque = "Doc n. " & TBContas!NDoctoBaixa
                                Case "TED": Cheque = "Ted n. " & TBContas!NDoctoBaixa
                                Case "MALOTE": Cheque = "Malote n. " & TBContas!NDoctoBaixa
                            End Select
                            Set TBFluxo = CreateObject("adodb.recordset")
                            If Left(TBContas!FormaBaixa, 6) = "CHEQUE" Or TBContas!FormaBaixa = "DOC" Or TBContas!FormaBaixa = "TED" Or TBContas!FormaBaixa = "MALOTE" Then
                                TextoFiltro = TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Crédito' and (idintconta = 0 or idintconta IS NULL)"
                            Else
                                If IsNull(TBContas!ID_varias) = True Or TBContas!ID_varias = 0 Then TextoFiltro = TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Crédito'" Else TextoFiltro = "ID_varias = " & TBContas!ID_varias
                            End If
                            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                TBFluxo!valor = Format(TBFluxo!valor - TBContas!valor, "###,##0.00")
                                TBFluxo.Update
                                If TBFluxo!valor <= 0 Then
                                    TBFluxo.Delete
                                    Conexao.Execute "DELETE from tbl_Contas_Varias where ID = " & IIf(IsNull(TBContas!ID_varias), 0, TBContas!ID_varias)
                                End If
                            End If
                        End If
                        
                        If TBContas!FormaBaixa <> "CHEQUE" And TBContas!FormaBaixa <> "CHEQUE PRÉ-DATADO" Then
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select * from tbl_instituicoes where txt_descricao = '" & TBContas!Banco & "' and ID_empresa = " & TBContas!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                TBProduto!Saldo = TBProduto!Saldo - TBContas!valor
                                TBProduto.Update
                            End If
                            TBProduto.Close
                        End If
                    End If
                    
                    Conexao.Execute "DELETE from familia_financeiro where Idconta = " & .ListItems(InitFor) & " and TipoConta = 'R' and Deposito_transf = 'False'"
                    Conexao.Execute "DELETE from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor)
                    
                    Conexao.Execute "DELETE from tbl_Detalhes_Recebimento where Idcontareceber = " & .ListItems(InitFor) & " and ID_nota = 0"
                    Conexao.Execute "DELETE from tbl_Detalhes_Recebimento_Nboletos where Idcontareceber = " & .ListItems(InitFor) & " and ID_nota = 0"
                    Conexao.Execute "UPDATE tbl_Detalhes_Recebimento Set Idcontareceber = 0 where Idcontareceber = " & .ListItems(InitFor)
                    Conexao.Execute "UPDATE tbl_Detalhes_Recebimento_Nboletos Set Idcontareceber = 0 where Idcontareceber = " & .ListItems(InitFor)
                    
                    Evento = "Excluir"
                    Documento1 = ""
                Else
                    IDlista = TBContas!IDDuplicata
                    Documento1 = "Duplicata: " & TBContas!IDDuplicata
                    TBContas!IDDuplicata = Null
                    TBContas.Update
                    
                    Evento = "Excluir duplicata"
                    
                    valor = 0
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select ISNULL(Sum(Valor), 0) as Valor from tbl_contas_receber where IDduplicata = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        valor = TBAbrir!valor
                    End If
                    TBAbrir.Close
                    If valor = 0 Then
                        Conexao.Execute "Delete from tbl_contas_receber_duplicatas where ID = " & IDlista
                    Else
                        NovoValor = Replace(valor, ",", ".")
                        Conexao.Execute "Update tbl_contas_receber_duplicatas Set Valor = " & NovoValor & " where ID = " & IDlista
                    End If
                End If
            End If
            TBContas.Close
            
            '==================================
            Modulo = "Financeiro/Contas a receber"
            ID_documento = .ListItems(InitFor)
            Documento = "Documento: " & .ListItems(InitFor).SubItems(5)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) " & TextoMsg & " antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox (TextoMsg1 & " excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    If Cmb_opcao_lista = "Excluir" Then ProcLimpaCampos
    ProcCarregaLista (1)
    Lista.SetFocus
    Novo_Receber = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmContas_receber_localizar.Show 1

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
ProcLimpaCampos
Novo_Receber = True
Chk_devolucao.Enabled = True
Chk_antecipacao.Enabled = True
Txt_data_transacao.SetFocus
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcParcelar()
On Error GoTo tratar_erro

If txtidintconta = "" Or txtValor.Text = "" Or txtIDcliente.Text = "" Then
    USMsgBox ("Informe a conta antes de parcelar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Receber = True Then
    USMsgBox ("Salve a conta antes de parcelar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Chk_antecipacao.Value = 1 Then
    USMsgBox ("Não é permitido parcelar esta conta, pois a mesma é uma antecipação."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Gerar_receb.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Receber = True Then
    If USMsgBox("A conta ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Receber = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Receber = False
ProcLimpaVariaveisCarregaLista
Imprimir = False
StrSql_Contas_Receber = ""
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
If txtidintconta = "" And Novo_Receber = False Then
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
If cmbtipo_conta = "" Then
    NomeCampo = "o tipo do documento"
    ProcVerificaAcao
    cmbtipo_conta.SetFocus
    Exit Sub
End If
txtparcela.PromptInclude = False
If Len(txtparcela) < 6 Then
    USMsgBox ("O número da parcela digitada não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
    txtparcela.SetFocus
    Exit Sub
End If
txtparcela.PromptInclude = True

valor = IIf(txtValor = "", 0, txtValor)
If Chk_devolucao.Value = 1 And valor >= 0 Then
    txtValor = Format(txtValor, "-###,##0.00")
ElseIf Chk_devolucao.Value = 0 And valor <= 0 Then
        NomeCampo = "o valor"
        ProcVerificaAcao
        txtValor.SetFocus
        Exit Sub
End If

If txtIDcliente.Text = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    cmdLocalizarCliente_Click
    Exit Sub
End If

'Verifica se é antecipação e se já foi vinculado em alguma conta paga
If Chk_antecipacao.Value = 1 Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_Contas_antecipacao where ID_antecipacao = " & IIf(txtidintconta = "", 0, txtidintconta) & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        USMsgBox ("Não é permitido salvar, pois esta antecipação já esta relacionada a uma conta baixada."), vbExclamation, "CAPRIND v5.0"
        TBContas.Close
        Exit Sub
    End If
End If

'Verifica se já existe conta com o mesmo número de nf e vencimento para o cliente
If Novo_Receber = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from tbl_contas_receber where NFiscal = '" & txtNFiscal & "' and Vencimento = '" & mskVencimento & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Antecipacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If TBProduto!IDCliente <> txtIDcliente Then
            USMsgBox ("Já existe uma conta cadastrada com este número de nota fiscal " & txtNFiscal & " com vencimento em " & mskVencimento & " para o cliente " & TBProduto!Nome_Razao & "."), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Já existe uma conta cadastrada com este número de nota fiscal " & txtNFiscal & " com vencimento em " & mskVencimento & " para este cliente."), vbExclamation, "CAPRIND v5.0"
        End If
        If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            TBProduto.Close
            Exit Sub
        End If
    End If
End If

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from tbl_contas_receber where IdIntConta = " & IIf(txtidintconta = "", 0, txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!status <> "TÍTULO EM ABERTO" And TBProduto!status <> "DUPLICATA DESCONTADA EM ABERTO" Then
        USMsgBox ("Não é permitido alterar esta conta, pois a mesma já foi baixada parcial, está bloqueada ou é uma antecipação baixada."), vbExclamation, "CAPRIND v5.0"
        TBProduto.Close
        Exit Sub
    End If
    'Corrige o valor das contas contábeis
    If TBProduto!valor <> valor And Lista_PC.ListItems.Count <> 0 Then
        If USMsgBox("Deseja atualizar o valor da(s) conta(s) contábil(eis)?)", vbYesNo) = vbYes Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Familia_financeiro where IDConta = " & IIf(txtidintconta = "", 0, txtidintconta) & " and TipoConta = 'R' and Valor > 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Valor1 = (TBAbrir!valor / TBProduto!valor) * 100
                    TBAbrir!valor = (valor * Valor1) / 100
                    TBAbrir.Update
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            ProcCarregaListaPC
        End If
    End If
'==============================================
' Verifica se tem duplicatas na NFe e edita
'==============================================
Set TBContas = CreateObject("adodb.recordset")
     TBContas.Open "Select * from tbl_detalhes_Recebimento where IDContaReceber = " & IIf(IsNull(TBProduto!IDintconta), 0, TBProduto!IDintconta), Conexao, adOpenKeyset, adLockOptimistic
      If TBContas.EOF = False Then
          TBContas!txt_Portador_Banco = cmbBanco
          TBContas!txt_Parcela = txtparcela
          TBContas!Tipo_doc = cmbtipo_conta
          TBContas!txt_tipopagto = cmb_forma
          TBContas!dt_Vencimento = mskVencimento.Value
          TBContas!Vencimento_boleto = mskVencimento.Value
         ' TBContas!dbl_valor = txtValor.Text
          TBContas!Valor_boleto = txtValor.Text
          TBContas.Update
          TBContas.Close
      End If
    
Else
    TBProduto.AddNew
    TBProduto!Parcial = False
    TBProduto!titulodesc = False
    TBProduto!Bloqueado = False
    TBProduto!Logsit = "N"
    TBProduto!IDtrocatitulo = 0
    TBProduto!Responsavel = pubUsuario
End If
If Chk_antecipacao.Value = 1 Then
    TBProduto!Antecipacao = True
    TBProduto!Saldo_antecipacao = txtValor.Text
Else
    TBProduto!Antecipacao = False
End If
If Chk_devolucao.Value = 1 Then TBProduto!Devolucao = True Else TBProduto!Devolucao = False

TBProduto!Data_transacao = Txt_data_transacao.Value
TBProduto!Tipo_doc = cmbtipo_conta
TBProduto!txt_ndocumento = txtDocumento
TBProduto!NFiscal = txtNFiscal.Text
TBProduto!Proposta = txtProposta.Text
TBProduto!emissao = mskEmissao.Value

If Cmb_tipo = "Cliente" Then
    Tipo = "CL"
    IDCli = txtIDcliente
ElseIf Cmb_tipo = "Fornecedor" Then
        Tipo = "FO"
        IDCli = txtIDcliente
    ElseIf Cmb_tipo = "Funcionário" Then
            Tipo = "FU"
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select ID from Funcionarios where Codigo = '" & txtIDcliente & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                IDCli = TBAbrir!ID
            End If
            TBAbrir.Close
        Else
            Tipo = "IN"
            IDCli = txtIDcliente
End If

TBProduto!Tipo = Tipo
TBProduto!Nome_Razao = txtNome_Razao.Text
TBProduto!IDCliente = IDCli
TBProduto!Cidade = IIf(txtCidade = "", Null, txtCidade)
TBProduto!Estado = IIf(cbo_UF = "", Null, cbo_UF)
TBProduto!valor = IIf(txtValor.Text <> "", txtValor.Text, "0")
TBProduto!ValorExtenso = FunValorExtenso(TBProduto!valor)
TBProduto!status = txtStatus
TBProduto!Vencimento = mskVencimento.Value
TBProduto!Parcela = txtparcela.Text
TBProduto!Observacoes = txtObservacao.Text
TBProduto!Banco = IIf(cmbBanco = "", Null, cmbBanco)
TBProduto!FormaBaixa = cmb_forma
TBProduto!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBProduto.Update
txtidintconta = TBProduto!IDintconta

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_Detalhes_Recebimento Where IDContaReceber = " & txtidintconta.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
TBContas!dt_Vencimento = mskVencimento.Value
TBContas.Update
End If
TBContas.Close

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBProduto!IDFluxo), 0, TBProduto!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = txtidintconta
TBFluxo!Operacao = "À Creditar"
TBFluxo!Data = mskVencimento
TBFluxo!valor = IIf(txtValor <> "", txtValor, "0")
TBFluxo!Descricao = txtNome_Razao
TBFluxo!status = "N"
TBFluxo!int_NotaFiscal = txtNFiscal
TBFluxo!Instituicao = IIf(cmbBanco = "", Null, cmbBanco)
TBFluxo!Documento = txtDocumento
TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "Update tbl_contas_receber set IDFLUXO = " & TBFluxo!IDFluxo & " where IdIntConta = " & txtidintconta
TBFluxo.Close

TBProduto.Close
If Novo_Receber = True Then
    USMsgBox ("Nova conta a receber cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    ProcConstruirFiltroPadrao "CR.IdIntConta = " & txtidintconta, "{tbl_contas_receber.IDintconta} = " & txtidintconta, True, True
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
1:
'==================================
Modulo = "Financeiro/Contas a receber"
ID_documento = txtidintconta
Documento = "Documento: " & txtDocumento
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Receber = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

Permitido1 = False
If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select CR.IdContaRecomprada, CR.Status, CR.Antecipacao, CR.IDintconta, CR.IDduplicata, CR.ID_empresa, ISNULL(I.Id, 0) as ID, CR.IDcliente, CR.Tipo from tbl_contas_receber CR LEFT JOIN tbl_Instituicoes I ON I.txt_Descricao = CR.Banco where CR.IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If Cmb_opcao_lista = "Excluir" Then
                        If IsNull(TBContas!IdContaRecomprada) = False And TBContas!IdContaRecomprada <> "" And TBContas!IdContaRecomprada <> "0" Then GoTo Proximo
                        If TBContas!status <> "TÍTULO EM ABERTO" And TBContas!status <> "BLOQUEADA" And TBContas!status <> "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                        If TBContas!Antecipacao = True Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select ID from tbl_contas_antecipacao where ID_antecipacao = " & TBContas!IDintconta & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                            TBAbrir.Close
                        End If
                    ElseIf Cmb_opcao_lista = "Status" Then
                            If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                        ElseIf Cmb_opcao_lista = "Baixar" Then
                                If TBContas!status = "BLOQUEADA" Or TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                            ElseIf Cmb_opcao_lista = "Gerar duplicata" Then
                                    Permitido1 = True
                                    If ClienteRec = "" Then ClienteRec = .ListItems.Item(InitFor).SubItems(7)
                                    If ClienteRec <> .ListItems.Item(InitFor).SubItems(7) Then GoTo Proximo
                                    
                                    If .ListItems.Item(InitFor).SubItems(13) <> "" Then
                                        If IDduplicataRec = 0 Then IDduplicataRec = .ListItems.Item(InitFor).SubItems(13)
                                        If IDduplicataRec <> .ListItems.Item(InitFor).SubItems(13) Then GoTo Proximo
                                    End If
                                ElseIf Cmb_opcao_lista = "Excluir duplicata" Then
                                        If IsNull(TBContas!IDDuplicata) = True Then GoTo Proximo
                                    ElseIf Cmb_opcao_lista = "Enviar boleto" Or Cmb_opcao_lista = "Gerar arquivo remessa" Then
                                            If ProcVerifDadosBoleto(TBContas!ID_empresa, TBContas!IDintconta, TBContas!ID, TBContas!IDCliente, TBContas!Tipo, "", False) = False Then GoTo Proximo
                    End If
                End If
                TBContas.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If
If Permitido1 = False Then ClienteRec = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Permitido1 = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select CR.IdContaRecomprada, CR.Status, CR.Antecipacao, CR.IDintconta, CR.IDduplicata, CR.ID_empresa, ISNULL(I.Id, 0) as ID, CR.IDcliente, CR.Tipo from tbl_contas_receber CR LEFT JOIN tbl_Instituicoes I ON I.txt_Descricao = CR.Banco where CR.IdIntConta = " & Item, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If Cmb_opcao_lista = "Excluir" Then
                    If IsNull(TBContas!IdContaRecomprada) = False And TBContas!IdContaRecomprada <> "" And TBContas!IdContaRecomprada <> "0" Then
                        USMsgBox ("Não é permitido excluir esta conta, pois a mesma está vinculada a uma recompra."), vbExclamation, "CAPRIND v5.0"
                        TBContas.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If TBContas!status <> "TÍTULO EM ABERTO" And TBContas!status <> "BLOQUEADA" And TBContas!status <> "TÍTULO LIQUIDADO ANTECIPADO" Then
                        USMsgBox ("Não é permitido excluir esta conta, pois a mesma já foi baixada parcial ou descontada."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If TBContas!Antecipacao = True Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ID from tbl_contas_antecipacao where ID_antecipacao = " & TBContas!IDintconta & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            USMsgBox ("Não é permitido excluir esta conta, pois a mesma é uma antecipação e já esta relacionada a uma conta baixada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            TBAbrir.Close
                            Exit Sub
                        End If
                        TBAbrir.Close
                    End If
                ElseIf Cmb_opcao_lista = "Status" Then
                        If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                            USMsgBox ("Não é permitido alterar o status desta conta, pois a mesma é uma antecipação líquidada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    ElseIf Cmb_opcao_lista = "Baixar" Then
                            If TBContas!status = "BLOQUEADA" Or TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                                USMsgBox ("Não é permitido baixar esta conta, pois a mesma está bloqueada ou é uma antecipação líquidada."), vbExclamation, "CAPRIND v5.0"
                                .ListItems.Item(InitFor).Checked = False
                                Exit Sub
                            End If
                            If TBContas!status = "TÍTULO EM ABERTO" Or TBContas!status = "DUPLICATA DESCONTADA EM ABERTO" Then
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select ID from Familia_financeiro where IDConta = " & TBContas!IDintconta & " and TipoConta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = True Then
                                    If USMsgBox("Esta conta não está amarrada em nenhuma conta contábil, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                                        .ListItems.Item(InitFor).Checked = False
                                        TBAbrir.Close
                                        Exit Sub
                                    End If
                                End If
                                TBAbrir.Close
                            End If
                        ElseIf Cmb_opcao_lista = "Gerar duplicata" Then
                                Permitido1 = True
                                If ClienteRec = "" Then ClienteRec = .ListItems.Item(InitFor).SubItems(7)
                                If ClienteRec <> .ListItems.Item(InitFor).SubItems(7) Then
                                    USMsgBox ("Só é permitido selecionar contas do cliente " & ClienteRec & "."), vbExclamation, "CAPRIND v5.0"
                                    .ListItems.Item(InitFor).Checked = False
                                    Exit Sub
                                End If
                                If .ListItems.Item(InitFor).SubItems(13) <> "" Then
                                    If IDduplicataRec = 0 Then IDduplicataRec = .ListItems.Item(InitFor).SubItems(13)
                                    If IDduplicataRec <> .ListItems.Item(InitFor).SubItems(13) Then
                                        USMsgBox ("Só é permitido selecionar contas vinculada a duplicata n. " & IDduplicataRec & "."), vbExclamation, "CAPRIND v5.0"
                                        .ListItems.Item(InitFor).Checked = False
                                        Exit Sub
                                    End If
                                End If
                            ElseIf Cmb_opcao_lista = "Excluir duplicata" Then
                                    If IsNull(TBContas!IDDuplicata) = True Then
                                        USMsgBox ("Não existe duplicata gerada para esta conta."), vbExclamation, "CAPRIND v5.0"
                                        .ListItems.Item(InitFor).Checked = False
                                        Exit Sub
                                    End If
                                ElseIf Cmb_opcao_lista = "Enviar boleto" Or Cmb_opcao_lista = "Gerar arquivo remessa" Then
                                        If ProcVerifDadosBoleto(TBContas!ID_empresa, TBContas!IDintconta, TBContas!ID, TBContas!IDCliente, TBContas!Tipo, IIf(Cmb_opcao_lista = "Enviar boleto", "enviar", "gerar arquivo remessa"), True) = False Then
                                            .ListItems.Item(InitFor).Checked = False
                                            Exit Sub
                                        End If
                End If
            End If
            TBContas.Close
        End If
    Next InitFor
End With
If Permitido1 = False Then
    IDduplicataRec = 0
    ClienteRec = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    ProcLiberaBotao
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBContas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBContas!ID_empresa) = False And TBContas!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBContas!ID_empresa
txtidintconta.Text = TBContas!IDintconta
Txt_data_transacao.Value = IIf(IsNull(TBContas!Data_transacao), Date, TBContas!Data_transacao)
txtDocumento = IIf(IsNull(TBContas!txt_ndocumento), "", TBContas!txt_ndocumento)
txtNFiscal.Text = IIf(IsNull(TBContas!NFiscal), "", TBContas!NFiscal)
txtStatus.Text = IIf(IsNull(TBContas!status), "", TBContas!status)
mskEmissao.Value = IIf(IsNull(TBContas!emissao), Date, Format(TBContas!emissao, "dd/mm/yyyy"))

'Verifica saldo da antecipação
If TBContas!Antecipacao = True Then qt = IIf(IsNull(TBContas!Saldo_antecipacao), 0, TBContas!Saldo_antecipacao) Else qt = IIf(IsNull(TBContas!valor), 0, TBContas!valor)
txtValor.Text = Format(qt, "###,##0.00")

mskVencimento.Value = IIf(IsNull(TBContas!Vencimento), Date, Format(TBContas!Vencimento, "dd/mm/yyyy"))
If TBContas!Antecipacao = True Then Chk_antecipacao.Value = 1 Else Chk_antecipacao.Value = 0
If TBContas!Devolucao = True Then Chk_devolucao.Value = 1 Else Chk_devolucao.Value = 0

If IsNull(TBContas!Parcela) = False And TBContas!Parcela <> "" Then txtparcela.Text = TBContas!Parcela

txtObservacao.Text = IIf(IsNull(TBContas!Observacoes), "", TBContas!Observacoes)

If TBContas!Tipo = "CL" Then
    Cmb_tipo = "Cliente"
    txtIDcliente = TBContas!IDCliente
ElseIf IsNull(TBContas!Tipo) = True Or TBContas!Tipo = "" Or TBContas!Tipo = "FO" Then
        Cmb_tipo = "Fornecedor"
        txtIDcliente = TBContas!IDCliente
    ElseIf TBContas!Tipo = "FU" Then
            Cmb_tipo = "Funcionário"
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select Codigo from Funcionarios where ID = " & TBContas!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                txtIDcliente = TBFornecedor!CODIGO
            End If
        Else
            Cmb_tipo = "Instituição bancária"
            txtIDcliente = TBContas!IDCliente
End If
txtCidade.Text = IIf(IsNull(TBContas!Cidade), "", TBContas!Cidade)
cbo_UF.Text = IIf(IsNull(TBContas!Estado), "", TBContas!Estado)
txtNome_Razao.Text = IIf(IsNull(TBContas!Nome_Razao), "", TBContas!Nome_Razao)

Chk_antecipacao.Enabled = True
Chk_devolucao.Enabled = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_contas_antecipacao where ID_conta = " & txtidintconta & " or ID_antecipacao = " & txtidintconta & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Or txtStatus = "TÍTULO LIQUIDADO ANTECIPADO" Then
    Chk_antecipacao.Enabled = False
    Chk_devolucao.Enabled = False
Else
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_contas_devolucao where ID_conta = " & txtidintconta & " or ID_devolucao = " & txtidintconta & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Chk_antecipacao.Enabled = False
        Chk_devolucao.Enabled = False
    End If
    TBAbrir.Close
End If

ProcCarregaProposta
Novo_Receber = False

NomeCampo = "o tipo do documento"
If IsNull(TBContas!Tipo_doc) = True Or TBContas!Tipo_doc = "" Then
cmbtipo_conta.ListIndex = 0
Else
cmbtipo_conta.Text = TBContas!Tipo_doc
End If

NomeCampo = "a forma da baixa prevista"
If IsNull(TBContas!FormaBaixa) = False And TBContas!FormaBaixa <> "" Then cmb_forma = TBContas!FormaBaixa
NomeCampo = "a instituição bancária prevista"
If IsNull(TBContas!Banco) = False And TBContas!Banco <> "" Then cmbBanco = TBContas!Banco
1:
    ProcCarregaListaPC

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        If NomeCampo = "a instituição bancária prevista" Then
            USMsgBox ("Não foi encontrado a instituição bancária prevista ou a mesma está bloqueada."), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Não foi encontrado " & NomeCampo & " desta conta."), vbExclamation, "CAPRIND v5.0"
        End If
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaPC()
On Error GoTo tratar_erro

If txtStatus <> "TÍTULO LIQUIDADO ANTECIPADO" Then TextoFiltro = "FF.Pago_recebido = 'False'" Else TextoFiltro = "(FF.Pago_recebido = 'True' or FF.Pago_recebido = 'False')"

Lista_PC.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select FF.ID, F.Codigo, F.txt_descricao, FF.Valor, CR.Antecipacao, CR.Saldo_antecipacao from (tbl_contas_receber CR INNER JOIN Familia_financeiro FF ON FF.IDConta = CR.IdIntConta) INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDConta = " & txtidintconta & " and FF.Tipoconta = 'R' and " & TextoFiltro & " and FF.Deposito_transf = 'False' order by F.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_PC.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            
            'Verifica saldo da antecipação
            'If TBLISTA!Antecipacao = True Then qt = IIf(IsNull(TBLISTA!Saldo_antecipacao), 0, TBLISTA!Saldo_antecipacao) Else qt = IIf(IsNull(TBLISTA!Valor), 0, TBLISTA!Valor)
            '.Item(.Count).SubItems(3) = Format(qt, "###,##0.00")
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), 0, Format(TBLISTA!valor, "###,##0.00"))
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_PC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_PC, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_PC_DblClick()
On Error GoTo tratar_erro

If Lista_PC.ListItems.Count = 0 Then Exit Sub
Qtde = 0
Valor_conta = ""

Mensagem:
    Valor_conta = InputBox("Favor informar o novo valor da conta contábil.")
    If Valor_conta = "" Then Exit Sub
    If IsNumeric(Valor_conta) = False Then
        USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    Qtde = Valor_conta
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_contas_receber where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If TBAbrir!Devolucao = True And Qtde >= 0 Then
            Qtde = Format(Qtde, "-###,##0.00")
        ElseIf TBAbrir!Devolucao = False And Qtde <= 0 Then
                USMsgBox ("Informe o novo valor antes de alterar."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
        End If
    End If
    
    'Verifica saldo das contas contábeis
    valor = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(Valor) as Valor from Familia_financeiro where IDConta = " & txtidintconta & " and TipoConta = 'R' and ID <> " & Lista_PC.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Valor, Devolucao from tbl_contas_receber where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qt = TBAbrir!valor
        Permitido = True
        If TBAbrir!Devolucao = True Then
            If (valor + Qtde) < qt Then Permitido = False
        Else
            If (valor + Qtde) > qt Then Permitido = False
        End If
        If Permitido = False Then
            USMsgBox ("Não é permitido alterar, pois o valor digitado ultrapassa o saldo da conta."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem
        End If
    End If
    TBAbrir.Close
    
    NovoValor = Replace(Qtde, ",", ".")
    Conexao.Execute "Update Familia_financeiro Set Valor = " & NovoValor & " where ID = " & Lista_PC.SelectedItem
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaPC
    '====================================
    Modulo = "Financeiro/Contas a receber"
    Evento = "Alterar valor da conta contábil"
    ID_documento = Lista_PC.SelectedItem
    Documento = "Documento: " & txtDocumento
    Documento1 = "Código do plano: " & Lista_PC.SelectedItem.ListSubItems(1) & " - Descrição do plano: " & Lista_PC.SelectedItem.ListSubItems(2)
    ProcGravaEvento
    '===================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarMes()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro.SelectedItem.key)
NomeRel = "Contas_receber.rpt"
If OptDomes.Value = True Then ProcConstruirFiltroPadrao "month(vencimento)= '" & M & "' and Year(vencimento) = '" & cmbAno & "'", "Month ({tbl_Contas_receber.vencimento}) = " & M & " and year ({tbl_Contas_receber.vencimento})= " & cmbAno, True, True
If OptAteomes.Value = True Then ProcConstruirFiltroPadrao "month(vencimento)<= '" & M & "' and Year(vencimento) = '" & cmbAno & "'", "Month ({tbl_Contas_receber.vencimento}) <= " & M & " and year ({tbl_Contas_receber.vencimento})= " & cmbAno, True, True
If TabFiltro.SelectedItem.key = "Vencidas" Then ProcConstruirFiltroPadrao "(vencimento) < '" & Date & "'", "{tbl_Contas_receber.vencimento} < Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")", True, True
ProcSalvarDadosRel False, False, False, Date, Date = False
StrSql_Contas_Receber_AntecTotal = ""
StrSql_Contas_Receber_DevTotal = ""
ProcCarregaLista (1)
Imprimir = True
Novo_Receber = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcConstruirFiltroPadrao(TextoFiltro As String, TextoFiltroRel As String, ApagarFiltroAntec As Boolean, ApagarFiltroDev As Boolean)
On Error GoTo tratar_erro

CamposFiltro = "CR.IDintconta, CR.emissao, CR.Vencimento, CR.Valor, CR.txt_ndocumento, CR.NFiscal, CR.Parcela, CR.Nome_Razao, CR.Responsavel, CR.ID_empresa, CR.IDduplicata, CR.Saldo_antecipacao, CR.Antecipacao"
If Left(TextoFiltro, 2) = "PN" Then INNERJOINPADRAO = " from tbl_contas_receber CR INNER JOIN tbl_proposta_nota PN ON PN.ID_nota = CR.ID_nota" Else INNERJOINPADRAO = " from tbl_contas_receber CR"
INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
INNERJOINTEXTOSUM = "Select SUM(CR.Valor) AS TotContas " & INNERJOINPADRAO
TextoFiltroPadrao = "CR.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and CR.logsit = 'N' and CR.bloqueado = 'False'" & IIf(ApagarFiltroAntec = True, " and CR.Antecipacao = 'False'", "") & IIf(ApagarFiltroDev = True, " and CR.Devolucao = 'False'", "")
TextoFiltroPadraoDESC = TextoFiltroPadrao & " and CR.status = 'DUPLICATA DESCONTADA EM ABERTO'"
TextoFiltroPadraoRel = "{tbl_contas_receber.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_contas_receber.LogSit} = 'N' and {tbl_contas_receber.bloqueado} = False" & IIf(ApagarFiltroAntec = True, " and {tbl_contas_receber.Antecipacao} = False", "") & IIf(ApagarFiltroDev = True, " and {tbl_contas_receber.Devolucao} = False", "")
OrdenarTexto = " group by " & CamposFiltro & " order by CR.vencimento, CR.IdIntConta"
StrSql_Contas_Receber = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
StrSql_Contas_ReceberTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadrao & " and CR.titulodesc = 'False'"
If ApagarFiltroAntec = True Then StrSql_Contas_Receber_AntecTotal = ""
If ApagarFiltroDev = True Then StrSql_Contas_Receber_DevTotal = ""
StrSql_Contas_ReceberDescTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDESC
FormulaRel_Contas_Receber = TextoFiltroRel & " and " & TextoFiltroPadraoRel

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDcliente_Change()
On Error GoTo tratar_erro

txtNome_Razao = ""
txtCidade = ""
cbo_UF = ""
If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    If Cmb_tipo = "Cliente" Then
        TBAbrir.Open "Select NomeRazao, Cidade, UF, Banco, Tipo_doc from Clientes where idcliente = " & txtIDcliente & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtNome_Razao = IIf(IsNull(TBAbrir!NomeRazao), "", TBAbrir!NomeRazao)
            txtCidade = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
            cbo_UF = IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
            NomeCampo = "instituição bancária prevista para recebimento do cliente."
            If IsNull(TBAbrir!Banco) = False And TBAbrir!Banco <> "" Then cmbBanco.Text = TBAbrir!Banco
            NomeCampo = "tipo do documento do cliente."
            If IsNull(TBAbrir!Tipo_doc) = False And TBAbrir!Tipo_doc <> "" Then cmbtipo_conta.Text = TBAbrir!Tipo_doc
        End If
    ElseIf Cmb_tipo = "Fornecedor" Then
            TBAbrir.Open "Select Nome_Razao, Cidade, Estado, Banco, Tipo_doc from compras_fornecedores where idcliente = " & txtIDcliente & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                txtNome_Razao = IIf(IsNull(TBAbrir!Nome_Razao), "", TBAbrir!Nome_Razao)
                txtCidade = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
                cbo_UF = IIf(IsNull(TBAbrir!Estado), "", TBAbrir!Estado)
                NomeCampo = "instituição bancária prevista para recebimento do fornecedor."
                If IsNull(TBAbrir!Banco) = False And TBAbrir!Banco <> "" Then cmbBanco.Text = TBAbrir!Banco
                NomeCampo = "tipo do documento do fornecedor."
                If IsNull(TBAbrir!Tipo_doc) = False And TBAbrir!Tipo_doc <> "" Then cmbtipo_conta.Text = TBAbrir!Tipo_doc
            End If
        ElseIf Cmb_tipo = "Funcionário" Then
                TBAbrir.Open "Select Nome from Funcionarios where Codigo = '" & txtIDcliente & "' and DtValidacao IS NOT NULL and Situacao = 'Normal'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtNome_Razao = TBAbrir!Nome
            Else
                TBAbrir.Open "Select Txt_descricao from tbl_Instituicoes where ID = " & txtIDcliente & " and DtValidacao IS NOT NULL and Bloqueado <> 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtNome_Razao = TBAbrir!Txt_descricao
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIdcliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: ProcLocalizarCliente
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNFiscal_LostFocus()
On Error GoTo tratar_erro

If txtNFiscal <> "" And IsNumeric(txtNFiscal) = True Then txtNFiscal = FunTamanhoTextoZeroEsq(txtNFiscal, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNome_Razao_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: ProcLocalizarCliente
End Select

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

Private Sub txtValor_Change()
On Error GoTo tratar_erro
    
If txtValor.Text <> "" Then
    VerifNumero = txtValor.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValor.Text = ""
        txtValor.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_LostFocus()
On Error GoTo tratar_erro
    
txtValor.Text = Format(txtValor.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaProposta()
On Error GoTo tratar_erro

With txtProposta
    .Clear
    
    If IsNull(TBContas!ID_nota) = False And TBContas!ID_nota <> "" And TBContas!ID_nota <> "0" Then
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select PN.Proposta from tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID where PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 1", Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Do While TBProposta.EOF = False
                If IsNull(TBProposta!Proposta) = False And TBProposta!Proposta <> "" Then .AddItem TBProposta!Proposta
                TBProposta.MoveNext
            Loop
        Else
            If IsNull(TBContas!Proposta) = False And TBContas!Proposta <> "" Then
                .AddItem TBContas!Proposta
                .Text = TBContas!Proposta
            End If
        End If
        TBProposta.Close
    Else
        If IsNull(TBContas!Proposta) = False And TBContas!Proposta <> "" Then
            .AddItem TBContas!Proposta
            .Text = TBContas!Proposta
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRelacionamento()
On Error GoTo tratar_erro

If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de visualizar a lista de contas relacionadas."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Financeiro_Contas_Pagas = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = True
Financeiro_Contas_Recebidas = False
frmContas_antecipacoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLiberaBotao()
On Error GoTo tratar_erro

With USToolBar1
    If TBContas!Antecipacao = True Then .ButtonState(14) = 0 Else .ButtonState(14) = 5
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarDadosRel(DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataInicio As Date, DataFinal As Date)
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = Formulario
    If DataTransacao = True Or DataEmissao = True Or DataVencimento = True Then
        TBLISTA!Data_inicial = DataInicio
        TBLISTA!Data_final = DataFinal
        TBLISTA!Turno = True
    Else
        TBLISTA!Turno = False
    End If
    TBLISTA.Update
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRetorno()
On Error GoTo tratar_erro

If txtidintconta = "" Or txtValor.Text = "" Or txtIDcliente.Text = "" Then
    USMsgBox ("Informe a conta antes de receber o arquivo retorno."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select CR.*, DR.Carteira, DR.Carteira1 from tbl_contas_receber CR INNER JOIN tbl_Detalhes_Recebimento DR ON CR.IDintconta = DR.IDContaReceber where CR.IDintconta = " & txtidintconta & " and Carteira is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Financeiro_Contas_Receber = True
    With CD1
        .filename = ""
        .Filter = "*.txt"
        .DialogTitle = "Carregar arquivo retorno"
        .Action = 1
        Contador = Len(.FileTitle)
        caminho = Left(.filename, Len(.filename) - Contador)
        FamiliaAntiga = .FileTitle
    End With
    Remessa = False
    Enviar_Email = False
    ProcPassaDadosContaCorrenteParaCobreBemX TBContas!Carteira, IIf(IsNull(TBContas!Carteira1), "", TBContas!Carteira1), "", Cmb_empresa.ItemData(Cmb_empresa.ListIndex), True, ""
    ProcBaixarRetorno caminho, FamiliaAntiga
    If Permitido = True Then
        USMsgBox ("Arquivo retorno recebido com sucesso."), vbInformation, "CAPRIND v5.0"
        ProcCarregaLista (1)
    Else
        USMsgBox ("Não foi possível fazer a leitura desse arquivo retorno, pois o layout do arquivo está diferente do padrão."), vbExclamation, "CAPRIND v5.0"
    End If
Else
    USMsgBox ("Não é possivel receber o arquivo retorno, pois esta conta não possui boleto emitido."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBaixarRetorno(DiretorioRet As String, ArquivoRet As String)
On Error GoTo tratar_erro

Permitido = False
CobreBemX1.OcorrenciasCobranca.Clear
CobreBemX1.ArquivoRetorno.Diretorio = DiretorioRet
CobreBemX1.ArquivoRetorno.Arquivo = ArquivoRet
CobreBemX1.CarregaArquivosRetorno
Dataini = 0
For i = 0 To CobreBemX1.OcorrenciasCobranca.Count - 1
    If CobreBemX1.OcorrenciasCobranca(i).Pagamento = True Then
        valor = CobreBemX1.OcorrenciasCobranca(i).ValorPago
        NovoValor = Replace(valor, ",", ".")
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select CR.* from tbl_contas_receber CR INNER JOIN tbl_Detalhes_Recebimento DR ON CR.IDintconta = DR.IDContaReceber where DR.Nosso_numero = '" & Mid(CobreBemX1.OcorrenciasCobranca(i).NossoNumero, 2, 9) & "' and CR.Valor <= " & NovoValor & " and CR.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Permitido = True
            
            'Data do recebimento
            DataFim = CobreBemX1.OcorrenciasCobranca(i).DataOcorrencia
            DataFim = DataFim + 1
            
            If Dataini = DataFim Then
                If ID_varias = 0 Then
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from tbl_Contas_Varias", Conexao, adOpenKeyset, adLockOptimistic
                    TBGravar.AddNew
                    TBGravar.Update
                    ID_varias = TBGravar!ID
                    TBGravar.Close
                End If
                
                Conexao.Execute "UPDATE tbl_contas_receber Set ID_varias = " & ID_varias & " where Banco = '" & TBAbrir!Banco & "' and Data_pagamento = '" & Dataini & "'"
                Conexao.Execute "UPDATE tbl_Fluxo_de_caixa Set Bloqueado = 'True' where Instituicao = '" & TBAbrir!Banco & "' and Data = '" & Dataini & "' and ID_varias = 0"
            Else
                ID_varias = 0
            End If
            
            If TBAbrir!Antecipacao = False Then
                TBAbrir!Logsit = "S"
                TBAbrir!valortitulorecebido = valor
                Valor1 = TBAbrir!valortitulorecebido
            Else
                TBAbrir!valortitulorecebido = 0
                Valor1 = TBAbrir!valor
            End If
        
            'TBAbrir!FormaBaixa = cmb_forma.Text
            TBAbrir!Data_pagamento = DataFim
            If TBAbrir!status = "TÍTULO RECEBIDO PARCIAL" Then
                TBAbrir!status = "TÍTULO RECEBIDO PARCIAL LIQUIDADO"
                TBAbrir!ValorPendente = 0
                TBAbrir!tituloref = TBAbrir!IDintconta
            Else
                If TBAbrir!status = "DUPLICATA DESCONTADA EM ABERTO" Then
                        TBAbrir!status = "DUPLICATA DESCONTADA LIQUIDADA"
                    ElseIf TBAbrir!Antecipacao = True Then
                            TBAbrir!status = "TÍTULO LIQUIDADO ANTECIPADO"
                        ElseIf TBAbrir!Devolucao = True Then
                                TBAbrir!status = "TÍTULO DEVOLVIDO LIQUIDADO"
                            Else
                                TBAbrir!status = "TÍTULO LIQUIDADO"
                End If
            End If
            If pubUsuario <> "" Then TBAbrir!resprec = pubUsuario
            'TBAbrir!NDoctoBaixa = txt_ndocumento.Text
            'TBAbrir!Banco = cmb_Banco.Text
            'If Contador > 1 Then TBAbrir!Obs = TBAbrir!Observacoes Else TBAbrir!Obs = txtObs.Text
            TBAbrir!Dias_atraso = IIf(TBAbrir!Data_pagamento - TBAbrir!Vencimento < 0, 0, TBAbrir!Data_pagamento - TBAbrir!Vencimento)
            
            'Família de contas
            Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'True' where IDConta = " & TBAbrir!IDintconta & " and tipoconta = 'R' and Pago_recebido = 'False'"
            
            TBAbrir!Juros_valor = CobreBemX1.OcorrenciasCobranca(i).ValorJurosPago
            TBAbrir!Multa_valor = CobreBemX1.OcorrenciasCobranca(i).ValorMultaPaga
            TBAbrir!Desconto_valor = CobreBemX1.OcorrenciasCobranca(i).ValorDesconto
            
            TBAbrir!ID_varias = ID_varias
            TBAbrir.Update
            
            'Fluxo de Caixa
            Set TBFluxo = CreateObject("adodb.recordset")
            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBAbrir!IDFluxo), 0, TBAbrir!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
            If TBFluxo.EOF = True Then TBFluxo.AddNew
            TBFluxo!IDintconta = TBAbrir!IDintconta
            TBFluxo!Operacao = "Crédito"
            TBFluxo!Data = TBAbrir!Data_pagamento
            TBFluxo!Descricao = TBAbrir!Nome_Razao
            TBFluxo!status = "S"
            TBFluxo!int_NotaFiscal = TBAbrir!NFiscal
            TBFluxo!Documento = TBAbrir!txt_ndocumento
            TBFluxo!Instituicao = TBAbrir!Banco
            TBFluxo!Hora = Format(Now, "hh:mm:ss")
            TBFluxo!Obs = TBFluxo!Descricao
            TBFluxo!Bloqueado = False
            TBFluxo!valor = Valor1
            If TBAbrir!titulodesc = True Then TBFluxo!Bloqueado = True
            TBFluxo!ID_empresa = TBAbrir!ID_empresa
            TBFluxo!ID_varias = 0
            'If txt_ndocumento <> "" Then TBFluxo!Cheque = txt_ndocumento
            TBFluxo!tituloref = TBAbrir!IDintconta
            If ID_varias <> 0 And TBAbrir!status <> "DUPLICATA DESCONTADA EM ABERTO" Then
                TBFluxo!Bloqueado = True
                
                'Cria registro com o valor total da operação
                valor = 0
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Sum(valortitulorecebido) as ValorTotal from tbl_contas_receber where ID_varias = " & ID_varias, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    valor = IIf(IsNull(TBFI!ValorTotal), 0, Format(TBFI!ValorTotal, "###,##0.00"))
                End If
                TBFI.Close
                
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from tbl_Fluxo_de_caixa where ID_varias = " & ID_varias, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then TBGravar.AddNew
                TBGravar!Operacao = "Crédito"
                TBGravar!Data = TBAbrir!Data_pagamento
                TBGravar!valor = valor
                TBGravar!Bloqueado = False
                TBGravar!Descricao = "RCBTO. VARIAS CONTAS"
                TBGravar!status = "S"
                TBGravar!Instituicao = TBAbrir!Banco
                TBGravar!Hora = TBFluxo!Hora
                'TBGravar!Cheque = txt_ndocumento
                TBGravar!Obs = TBGravar!Descricao
                TBGravar!ID_empresa = TBAbrir!ID_empresa
                TBGravar!ID_varias = ID_varias
                TBGravar.Update
                TBGravar.Close
            End If
            TBFluxo.Update
            Conexao.Execute "Update tbl_contas_receber set IDFLUXO = " & TBFluxo!IDFluxo & " where IDIntconta = " & TBAbrir!IDintconta
            TBFluxo.Close
            
            If TBAbrir!status <> "DUPLICATA DESCONTADA LIQUIDADA" Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from tbl_instituicoes where txt_Descricao = '" & TBAbrir!Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    TBItem!Saldo = Format(TBItem!Saldo + Valor1, "###,##0.00")
                    TBItem.Update
                End If
                TBItem.Close
            End If
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select local_troca from Troca_titulo where ID = " & IIf(IsNull(TBAbrir!IDtrocatitulo), 0, TBAbrir!IDtrocatitulo), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                Set TBReceber = CreateObject("adodb.recordset")
                TBReceber.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBContas!local_troca & "' and tbl_contas_receber.ID_empresa = " & TBAbrir!ID_empresa & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
                If TBReceber.EOF = False Then
                    valor = IIf(IsNull(TBReceber!valor), 0, TBReceber!valor)
                    NovoValor = Replace(valor, ",", ".")
                    Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & TBContas!local_troca & "' and ID_empresa = " & TBAbrir!ID_empresa
                End If
                TBReceber.Close
            End If
            TBContas.Close
            
            '==================================
            Modulo = "Financeiro/Contas a receber"
            Evento = "Baixar conta"
            ID_documento = TBAbrir!IDintconta
            Documento = "Documento: " & TBAbrir!txt_ndocumento
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
            Dataini = TBAbrir!Data_pagamento
        End If
        TBAbrir.Close
    End If
Next i

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcPlanoContas
    Case 7: ProcAgendaDia
    Case 8: ProcParcelar
    Case 9: ProcCopiar
    Case 10: ProcReceber
    Case 11: ProcRecomprar
    Case 12: ProcEmitirBoleto
    Case 13: ProcStatus
    Case 14: ProcRelacionamento
    Case 15: ProcRetorno
    Case 16: ProcAtualizar
    Case 18: ProcAjuda
    Case 19: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
