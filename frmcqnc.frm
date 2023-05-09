VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcqnc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Não conformidade"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
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
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView Listafases 
      Height          =   4665
      Left            =   60
      TabIndex        =   23
      Top             =   4500
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8229
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "ID_OS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qtde. NC"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "QT.Cond."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Operador"
         Object.Width           =   3795
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Disposição"
         Object.Width           =   5912
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Disposição*"
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
      Height          =   555
      Left            =   60
      TabIndex        =   71
      Top             =   3930
      Width           =   15195
      Begin VB.OptionButton Opt_nada_consta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nada consta"
         DisabledPicture =   "frmcqnc.frx":0000
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
         TabIndex        =   80
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Opt_aprovado_desvio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aprovado com desvio"
         DisabledPicture =   "frmcqnc.frx":249F42
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   1425
         TabIndex        =   79
         Top             =   300
         Width           =   1845
      End
      Begin VB.OptionButton Opt_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outros"
         DisabledPicture =   "frmcqnc.frx":493E84
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
         Left            =   12660
         TabIndex        =   78
         Top             =   300
         Width           =   795
      End
      Begin VB.OptionButton Opt_reaproveitar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reaproveitar para outro produto"
         DisabledPicture =   "frmcqnc.frx":6DDDC6
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
         Left            =   6960
         TabIndex        =   77
         Top             =   300
         Width           =   2685
      End
      Begin VB.OptionButton Opt_aprovado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aprovado"
         DisabledPicture =   "frmcqnc.frx":927D08
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   270
         TabIndex        =   76
         Top             =   300
         Width           =   1005
      End
      Begin VB.OptionButton Opt_selecionar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Selecionar"
         DisabledPicture =   "frmcqnc.frx":B71C4A
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
         Height          =   210
         Left            =   5775
         TabIndex        =   75
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton Opt_retrabalhar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Retrabalhar"
         DisabledPicture =   "frmcqnc.frx":DBBB8C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   210
         Left            =   4470
         TabIndex        =   74
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton Opt_rejeitar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rejeitar"
         DisabledPicture =   "frmcqnc.frx":1005ACE
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
         Height          =   210
         Left            =   3420
         TabIndex        =   73
         Top             =   300
         Width           =   885
      End
      Begin VB.OptionButton optDevolver 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Devolver para fornecedor/cliente"
         DisabledPicture =   "frmcqnc.frx":124FA10
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
         Left            =   9795
         TabIndex        =   72
         Top             =   300
         Width           =   2715
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
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   54
      Top             =   9120
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
         ItemData        =   "frmcqnc.frx":1499952
         Left            =   7080
         List            =   "frmcqnc.frx":149995C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         TabStop         =   0   'False
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
         Left            =   3060
         TabIndex        =   24
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
         TabIndex        =   25
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   29
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":1499971
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
         TabIndex        =   28
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":149D118
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
         TabIndex        =   26
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
         TabIndex        =   27
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":14A0C27
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
         TabIndex        =   30
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":14A4D1B
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
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
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
         Index           =   29
         Left            =   5790
         TabIndex        =   66
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
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
         Left            =   3690
         TabIndex        =   64
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
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
         Left            =   2370
         TabIndex        =   57
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
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
         Left            =   180
         TabIndex        =   56
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
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
         Left            =   13050
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   55
      TabIndex        =   31
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   8160
         TabIndex        =   85
         ToolTipText     =   "Quantidade da ordem"
         Top             =   945
         Width           =   855
      End
      Begin VB.Frame Frame5 
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
         Height          =   675
         Left            =   0
         TabIndex        =   81
         Top             =   0
         Width           =   1545
         Begin VB.CheckBox chkAnalizada 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Analisada"
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
            Left            =   210
            TabIndex        =   82
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtid 
         Height          =   375
         Left            =   15240
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txtQTCD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9900
         TabIndex        =   13
         ToolTipText     =   "Quantidade de unidades aprovadas com desvio"
         Top             =   945
         Width           =   915
      End
      Begin VB.ComboBox Cmb_maquina 
         Appearance      =   0  'Flat
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
         Left            =   4005
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   61
         ToolTipText     =   "Posto de trabalho."
         Top             =   345
         Width           =   1215
      End
      Begin VB.TextBox txtID_SD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   13950
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "N° da solicitação de desvio."
         Top             =   945
         Width           =   735
      End
      Begin VB.CommandButton cmdSD 
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
         Left            =   14700
         Picture         =   "frmcqnc.frx":14A85A8
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Criar SD."
         Top             =   945
         Width           =   315
      End
      Begin VB.TextBox txtidos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "ID."
         Top             =   345
         Width           =   810
      End
      Begin VB.TextBox txtpubusuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   11865
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   345
         Width           =   3150
      End
      Begin VB.TextBox Txt_turno 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Turno."
         Top             =   345
         Width           =   705
      End
      Begin VB.CommandButton cmdRNC 
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
         Left            =   13620
         Picture         =   "frmcqnc.frx":14A868A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Visualizar RNC."
         Top             =   945
         Width           =   315
      End
      Begin VB.TextBox txtReferencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Código de referência."
         Top             =   945
         Width           =   1455
      End
      Begin VB.ComboBox cmbOperador 
         Appearance      =   0  'Flat
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
         Left            =   8820
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Operador."
         Top             =   345
         Width           =   3045
      End
      Begin VB.ComboBox cmbOS 
         Appearance      =   0  'Flat
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
         Left            =   3150
         Sorted          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Número da OS."
         Top             =   345
         Width           =   840
      End
      Begin VB.TextBox txtDesenho 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox txtDescricao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   945
         Width           =   5205
      End
      Begin VB.TextBox TxtRNC 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "N° da RNC."
         Top             =   945
         Width           =   885
      End
      Begin VB.TextBox txtfase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Fase."
         Top             =   345
         Width           =   585
      End
      Begin VB.TextBox txtnc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9030
         TabIndex        =   12
         ToolTipText     =   "Quantidade de unidades não conforme."
         Top             =   945
         Width           =   855
      End
      Begin VB.TextBox txtaprovadas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   10830
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de unidadades aprovadas"
         Top             =   945
         Width           =   885
      End
      Begin VB.TextBox txtlote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de unidades produzidas"
         Top             =   945
         Width           =   975
      End
      Begin VB.TextBox txtof 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2385
         TabIndex        =   1
         ToolTipText     =   "Número da ordem."
         Top             =   345
         Width           =   750
      End
      Begin MSComCtl2.DTPicker txtdata 
         Height          =   315
         Left            =   5835
         TabIndex        =   4
         ToolTipText     =   "Data."
         Top             =   345
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   198574083
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txthora 
         Height          =   315
         Left            =   7035
         TabIndex        =   5
         ToolTipText     =   "Hora."
         Top             =   345
         Width           =   1035
         _ExtentX        =   1826
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
         Format          =   198574082
         CurrentDate     =   39055
      End
      Begin VB.TextBox Txt_ID_RNC 
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "ID RNC."
         Top             =   960
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   8325
         TabIndex        =   86
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Com desvio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9945
         TabIndex        =   67
         Top             =   750
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Posto trabalho*"
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
         Left            =   4035
         TabIndex        =   62
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº SD"
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
         Left            =   14100
         TabIndex        =   60
         Top             =   750
         Width           =   420
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   1875
         TabIndex        =   53
         Top             =   150
         Width           =   195
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
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
         Left            =   12825
         TabIndex        =   50
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Turno*"
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
         Left            =   8205
         TabIndex        =   49
         Top             =   150
         Width           =   510
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   255
         TabIndex        =   46
         Top             =   750
         Width           =   1050
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código referência"
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
         Left            =   1590
         TabIndex        =   45
         Top             =   750
         Width           =   1275
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   5512
         TabIndex        =   44
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº RNC"
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
         Left            =   12885
         TabIndex        =   43
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Não conf."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   9075
         TabIndex        =   40
         Top             =   750
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aprovado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   10905
         TabIndex        =   39
         Top             =   750
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora*"
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
         Left            =   7335
         TabIndex        =   38
         Top             =   150
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fase*"
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
         Left            =   5370
         TabIndex        =   37
         Top             =   150
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data*"
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
         Left            =   6210
         TabIndex        =   36
         Top             =   150
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "OS*"
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
         Left            =   3390
         TabIndex        =   35
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   12015
         TabIndex        =   34
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Operador*"
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
         Left            =   9945
         TabIndex        =   33
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem*"
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
         Left            =   2535
         TabIndex        =   32
         Top             =   150
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1665
      Left            =   55
      TabIndex        =   41
      Top             =   2280
      Width           =   15195
      Begin DrawSuite2022.USButton btnNS 
         Height          =   735
         Left            =   13680
         TabIndex        =   83
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         DibPicture      =   "frmcqnc.frx":14A876C
         Caption         =   "Número de série"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.TextBox TxtdescricaoNC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Descrição da não conformidade."
         Top             =   390
         Width           =   5715
      End
      Begin VB.ComboBox Cmb_origem 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Origem."
         Top             =   390
         Width           =   2325
      End
      Begin VB.TextBox txtParecerF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2970
         TabIndex        =   20
         ToolTipText     =   "Dimensão."
         Top             =   390
         Width           =   5565
      End
      Begin VB.TextBox txtobscq 
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
         Height          =   585
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         ToolTipText     =   "Observação do parecer do controle de qualidade."
         Top             =   960
         Width           =   13395
      End
      Begin DrawSuite2022.USButton Cmd_cadastrar_origem 
         Height          =   315
         Left            =   2520
         TabIndex        =   68
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   390
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":14B2219
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
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   0
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton Cmd_cadastrar_causa 
         Height          =   315
         Left            =   14700
         TabIndex        =   69
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   390
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":14D031E
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
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   0
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton cmdPFP 
         Height          =   315
         Left            =   8550
         TabIndex        =   70
         ToolTipText     =   "Buscar item cadastrado ou vendido"
         Top             =   390
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmcqnc.frx":14EE423
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
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   0
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Origem"
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
         Left            =   1087
         TabIndex        =   52
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
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
         Left            =   6405
         TabIndex        =   48
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição da não conformidade"
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
         Left            =   10702
         TabIndex        =   47
         Top             =   180
         Width           =   2250
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensão"
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
         Left            =   5407
         TabIndex        =   42
         Top             =   180
         Width           =   690
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   51
      Top             =   0
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
      ButtonCaption6  =   "Rep./Retrab."
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Visualizar reposição/retrabalho (F8)"
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
      ButtonWidth6    =   81
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "OS retrabalho"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Criar OS de retrabalho (F9)"
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
      ButtonLeft7     =   322
      ButtonTop7      =   2
      ButtonWidth7    =   75
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Ordem retrabalho"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Criar ordem de retrabalho"
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
      ButtonLeft8     =   399
      ButtonTop8      =   2
      ButtonWidth8    =   93
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "RNC"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Criar a RNC (F10)"
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
      ButtonLeft9     =   494
      ButtonTop9      =   2
      ButtonWidth9    =   29
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Disposição"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Informar disposição (F11)"
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
      ButtonLeft10    =   525
      ButtonTop10     =   2
      ButtonWidth10   =   58
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Atualizar"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft11    =   585
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
      ButtonLeft12    =   646
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft13    =   650
      ButtonTop13     =   2
      ButtonWidth13   =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft14    =   693
      ButtonTop14     =   2
      ButtonWidth14   =   30
      ButtonHeight14  =   21
      ButtonUseMaskColor14=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13890
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmcqnc.frx":150C528
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   63
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
Attribute VB_Name = "frmcqnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_CQNC           As Boolean
Public pesquisaPorOrdem    As Boolean
Public StrSql_CQ_NC        As String
Public StrSql_CQ_NC_FIltro As String
Public FormulaRel_CQ_NC    As String
Public ControleNC          As String
Public dataTipo_CQNC_DE    As String
Public dataTipo_CQNC_Ate   As String
Public dataTipo_CQNC       As Integer
Dim TBLISTA_NC             As ADODB.Recordset
Dim MesmaOrdem             As Long
Dim MesmaOS                As Long
Dim MesmoItem              As String
Dim OrdemRetrabalho        As Long
Dim QtdeOrdemRetrabalho    As Double
Dim PrazoEntregaOrdem      As Date

Private Sub btnNS_Click()
On Error GoTo tratar_erro

frmNumeroSerie.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With ListaFases
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        If PCP_Ordem = False Then .ButtonState(4) = 0 Else .ButtonState(4) = 5
        .ButtonState(8) = 5
        .ButtonState(9) = 5
        .ButtonState(10) = 5
    ElseIf Cmb_opcao_lista = "RNC" Then
        .ButtonState(4) = 5
        .ButtonState(8) = 5
        .ButtonState(9) = 0
        .ButtonState(10) = 5
    ElseIf Cmb_opcao_lista = "Ordem de retrabalho" Then
        .ButtonState(4) = 5
        .ButtonState(8) = 0
        .ButtonState(9) = 5
        .ButtonState(10) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(8) = 5
        .ButtonState(9) = 5
        .ButtonState(10) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbOS_Change()
On Error GoTo tratar_erro

If cmbOS.Text <> "" Then
    VerifNumero = cmbOS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        cmbOS.Text = ""
        cmbOS.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbOS_LostFocus()
On Error GoTo tratar_erro

ProcLimpaCamposOS
Set TBOrdem = CreateObject("adodb.recordset")
'TBOrdem.Open "Select OS.Ordem, OS.Maquina, OS.Fase, OS.Quantidade, OS.QTOK, P.Desenho, P.N_Referencia, P.Produto from ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.idproducao = " & IIf(cmbOS = "", 0, cmbOS), Conexao, adOpenKeyset, adLockOptimistic
TBOrdem.Open "Select OS.Ordem, OS.Maquina, OS.Fase, OS.TotalProd as Quantidade, OS.QTOK, P.Desenho, P.N_Referencia, P.Produto from ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.idproducao = " & IIf(cmbOS = "", 0, cmbOS), Conexao, adOpenKeyset, adLockOptimistic

If TBOrdem.EOF = False Then
    txtof = IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
    If IsNull(TBOrdem!maquina) = False And TBOrdem!maquina <> "" Then Cmb_maquina = TBOrdem!maquina
    txtFase.Text = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
    txtdesenho.Text = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
    txtreferencia = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
    txtdescricao.Text = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
    txtLote.Text = IIf(IsNull(TBOrdem!quantidade), "", TBOrdem!quantidade)
    txtaprovadas.Text = IIf(IsNull(TBOrdem!QTOK), 0, TBOrdem!QTOK)
End If
TBOrdem.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLimpaCamposOS()
On Error GoTo tratar_erro

txtof.Text = ""
Cmb_maquina.ListIndex = -1
txtFase.Text = ""
txtdesenho.Text = ""
txtreferencia = ""
txtdescricao.Text = ""
txtLote.Text = ""
txtaprovadas.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmcqnc_abrir.Show 1
Frame3.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

'If InputBox("Informe a senha para liberar.") = "280362N" Then 'Then frmcqnc_atualizar.Show 1
'StrSql = ""

Set TBOrdem = CreateObject("adodb.recordset")
StrSql = "select IDProducao, Ordem, OS, reprovada, Quant, Data, maquina,Turno,Usuario, QTCD from ProducaoFases PFB where QTCD > '0' and Reprovada='0'"
TBOrdem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
TBOrdem.MoveFirst
Do While TBOrdem.EOF = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_NC_FABRICA WHERE IDProducao = " & TBOrdem!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
TBAbrir.AddNew
End If

TBAbrir!IDProducao = TBOrdem!IDProducao
TBAbrir!Ordem = TBOrdem!Ordem
TBAbrir!OS = TBOrdem!OS
TBAbrir!TTNC = TBOrdem!Reprovada
TBAbrir!LOTE = TBOrdem!Quant
TBAbrir!Data = TBOrdem!Data
TBAbrir!Hora = Hour(TBOrdem!Data)
TBAbrir!maquina = TBOrdem!maquina
TBAbrir!Turno = TBOrdem!Turno
TBAbrir!ParecerCQ = "Aprovado c/ desvio"
TBAbrir!Operador = TBOrdem!Usuario
TBAbrir!Setor = "QUALIDADE"
TBAbrir!obsFab = "5BASE"
TBAbrir!Analizada = True
TBAbrir!QTCD = TBOrdem!QTCD
TBAbrir.Update
TBOrdem.MoveNext
Loop
End If
TBAbrir.Close
TBOrdem.Close

'====================================================================================================

Set TBOrdem = CreateObject("adodb.recordset")
StrSql = "select PFB.IDProducao, PFB.Ordem, OS, reprovada, Quant, Data, PFB.maquina,Turno,Usuario, OS.QTCD from ProducaoFases_Backup PFB LEFT OUTER JOIN Ordemservico OS on  PFB.os = OS.IDProducao where os.QTCD > '0' and PFB.Reprovada='0' AND PFB.Descricao = 'FIM DE PRODUÇÃO'"
TBOrdem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
TBOrdem.MoveFirst
Do While TBOrdem.EOF = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_NC_FABRICA WHERE IDProducao = " & TBOrdem!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
TBAbrir.AddNew
End If

TBAbrir!IDProducao = TBOrdem!IDProducao
TBAbrir!Ordem = TBOrdem!Ordem
TBAbrir!OS = TBOrdem!OS
TBAbrir!TTNC = TBOrdem!Reprovada
TBAbrir!LOTE = TBOrdem!Quant
TBAbrir!Data = TBOrdem!Data
TBAbrir!Hora = Hour(TBOrdem!Data)
TBAbrir!maquina = TBOrdem!maquina
TBAbrir!Turno = TBOrdem!Turno
TBAbrir!ParecerCQ = "Aprovado c/ desvio"
TBAbrir!Operador = TBOrdem!Usuario
TBAbrir!Setor = "QUALIDADE"
TBAbrir!obsFab = "5BASE"
TBAbrir!Analizada = True
TBAbrir!QTCD = TBOrdem!QTCD
TBAbrir.Update
TBOrdem.MoveNext
Loop
End If
TBAbrir.Close
TBOrdem.Close

'========================================================================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmcqnc_atualizar
        If .Chk1.Value = 1 Then
            'Atualizar ID do apontamento e dados da(s) NC
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from CQ_NC_FABRICA order by OS", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select * FROM producao INNER JOIN Ordemservico ON Ordemservico.Ordem = producao.Ordem where Ordemservico.IDProducao = " & TBproducao!OS & " and Producao.Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        NomeTabelaAp = "ProducaoFases_Backup"
                    Else
                        NomeTabelaAp = "ProducaoFases"
                    End If
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select IDproducao, CodigoDesc from " & NomeTabelaAp & " where OS = " & TBproducao!OS & " order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        TBAfericao.Find ("Idproducao = " & TBproducao!IDProducao)
                        TBAfericao.MovePrevious
                        If TBAfericao.BOF = False Then
                            If TBAfericao!CodigoDesc = 2 Then
                                TBproducao!IDProducao = TBAfericao!IDProducao
                                TBproducao.Update
                            End If
                        End If
                    End If
                    TBAfericao.Close
                    TBproducao.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from producaofases order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                    ProcCriaNCAtualizacao
                    TBproducao.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from producaofases_Backup order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                    ProcCriaNCAtualizacao
                    TBproducao.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
        
        If .Chk2.Value = 1 Then
            'Atualiza quantidade(s) da(s) NC na(s) Ordem(ns) e OS('s)
            OS = 0
            QTNC = 0
            OF = 0
            TotalNC = 0
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from CQ_NC_FABRICA order by OS", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                    QTNC = QTNC + IIf(IsNull(TBproducao!TTNC), 0, TBproducao!TTNC)
                    OS = TBproducao!OS
                    TBproducao.MoveNext
                    If TBproducao.EOF = False Then
                        If TBproducao!OS <> OS Then
                            Set TBGravar = CreateObject("adodb.recordset")
                            TBGravar.Open "Select * from Ordemservico where IDProducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
                            If TBGravar.EOF = False Then
                                TBGravar!QTNC = QTNC
                                TBGravar!Totalprod = TBGravar!QTOK + QTNC
                                TBGravar.Update
                            End If
                            TBGravar.Close
                            QTNC = 0
                        End If
                    End If
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from ordemservico order by Ordem, idproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBproducao!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        TotalNC = TotalNC + TBproducao!QTNC
                    End If
                    TBAbrir.Close
                    OF = TBproducao!Ordem
                    TBproducao.MoveNext
                    If TBproducao.EOF = False Then
                        If TBproducao!Ordem <> OF Then
                            Conexao.Execute "Update Producao Set quantNC = '" & TotalNC & "' where Ordem = " & OF
                            TotalNC = 0
                        End If
                    End If
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBproducao.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Atualizar status da(s) NC
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from CQ_NC_FABRICA order by OS", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                    If IsNull(TBproducao!ParecerCQ) = True Or TBproducao!ParecerCQ = "" Then TBproducao!ParecerCQ = "Nada consta"
                    TBproducao.Update
                    TBproducao.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBproducao.Close
            Conexao.Execute "Update CQ_NC_FABRICA Set PARECERCQ = 'Aprovado' where PARECERCQ = 'APROVADO PELO C.Q.'"
            Conexao.Execute "Update CQ_NC_FABRICA Set PARECERCQ = 'Rejeitar' where PARECERCQ = 'REJEITAR'"
            Conexao.Execute "Update CQ_NC_FABRICA Set PARECERCQ = 'Retrabalhar' where PARECERCQ = 'RETRABALHAR'"
            Conexao.Execute "Update CQ_NC_FABRICA Set PARECERCQ = 'Selecionar' where PARECERCQ = 'SELECIONAR'"
            Conexao.Execute "Update CQ_NC_FABRICA Set PARECERCQ = 'Reaproveitar' where PARECERCQ = 'REAPROVEITAR PARA OUTRO PRODUTO.'"
        End If
        If .Chk4.Value = 1 Then
            'Atualizar Ordens
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "select data, Ordem, OS from CQ_NC_FABRICA where data >'08-10-2019'", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBproducao.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBproducao.EOF = False
                        Set TBLISTA = CreateObject("adodb.recordset")
                        TBLISTA.Open "select Ordem from OrdemServico where IDProducao = '" & TBproducao!OS & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBLISTA.EOF = False Then
                            TBproducao!Ordem = TBLISTA!Ordem
                            End If
                            
                    TBproducao.Update
                    TBproducao.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBproducao.Close
        End If
        
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Não conformidade"
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
End Sub

Private Sub ProcCriaNCAtualizacao()
On Error GoTo tratar_erro

If TBproducao!Reprovada <> 0 Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from CQ_NC_FABRICA where IDProducao = " & TBproducao!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
    TBGravar!IDProducao = TBproducao!IDProducao
    TBGravar!Ordem = TBproducao!Ordem
    TBGravar!OS = TBproducao!IDFase
    TBGravar!TTNC = TBproducao!Reprovada
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from Usuarios where Usuario = '" & TBproducao!Usuario & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        TBGravar!Operador = TBUsuarios!CODIGO & "-" & TBproducao!Usuario
    End If
    TBUsuarios.Close
    TBGravar!LOTE = TBproducao!Quant
    TBGravar!Data = Format(TBproducao!Data, "dd/mm/yy")
    TBGravar!Hora = Format(TBproducao!TempoInicio, "hh:mm:ss")
    TBGravar.Update
    TBGravar.Close
Else
    Conexao.Execute "DELETE from CQ_NC_FABRICA where IDProducao = " & TBproducao!IDProducao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) não conformidade(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select CQNCF.*, P.Desenho from CQ_NC_FABRICA CQNCF INNER JOIN Producao P ON CQNCF.Ordem = P.Ordem where CQNCF.Codigo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                ProcExcluirNCOSMaq TBFI!OS
                Conexao.Execute "DELETE from cq_nc_fabrica where Codigo = " & .ListItems(InitFor)
                If IsNull(TBFI!ID_RNC) = False Then Conexao.Execute "DELETE from CQ_RNC where ID = " & TBFI!ID_RNC
                If IsNull(TBFI!ID_SD) = False Then Conexao.Execute "DELETE from CQ_SD where ID = " & TBFI!ID_SD
                Conexao.Execute "DELETE from OrdemServico where Ordem = " & TBFI!Ordem & " and fase = " & .ListItems(InitFor).ListSubItems(5) & " and retrabalho = 'True'"
                
                '==================================
                Modulo = "Qualidade/Não conformidade"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Data: " & Format(TBFI!Data, "dd/mm/yy") & " - Hora: " & Format(TBFI!Hora, "hh:mm:ss") & " - Ordem: " & TBFI!Ordem & " - OS: " & TBFI!OS & " - Cód. interno: " & TBFI!Desenho & " - Operador: " & TBFI!Operador
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) não conformidade(s) antes de excluir ou alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Não conformidade(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame1.Enabled = False
    Frame3.Enabled = False
    Novo_CQNC = False
    ProcCarregaLista (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "CQ_nc.rpt"

ProcVerifRelPersonalizado
If PermitidoRel = False Or Left(Nome_banco, 9) <> "PROJETAR" And Right(Nome_banco, 9) <> "PROJETAR" Then ProcImprimirRel FormulaRel_CQ_NC, "" Else ProcImprimirRel_Variavel FormulaRel_CQ_NC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos
Novo_CQNC = True
Frame1.Enabled = True
Frame3.Enabled = True
txtobscq.Enabled = True
cmbOperador.Locked = False
txtnc.Locked = False
txtof.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_cadastrar_causa_Click()
On Error GoTo tratar_erro

Sit_REG = 0
frmcqnc_causa.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmd_cadastrar_origem_Click()
On Error GoTo tratar_erro

Sit_REG = 0
frmcqnc_origem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_NC.AbsolutePage <> 2 Then
    If TBLISTA_NC.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_NC.PageCount - 1)
    Else
        TBLISTA_NC.AbsolutePage = TBLISTA_NC.AbsolutePage - 2
        ProcExibePagina (TBLISTA_NC.AbsolutePage)
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
    TBLISTA_NC.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_NC.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_NC.AbsolutePage = 1
ProcExibePagina (TBLISTA_NC.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_NC.AbsolutePage <> -3 Then
    If TBLISTA_NC.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_NC.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_NC.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_NC.AbsolutePage = TBLISTA_NC.PageCount
ProcExibePagina (TBLISTA_NC.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPFP_Click()
On Error GoTo tratar_erro

Acao = "localizar as dimensões"
If txtof = "" Then
    NomeCampo = "a ordem"
    ProcVerificaAcao
    txtof.SetFocus
    Exit Sub
End If
If cmbOS = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    cmbOS.SetFocus
    Exit Sub
End If
Sit_REG = 0
frmcqnc_dimensoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdRNC_Click()
On Error GoTo tratar_erro

Acao = "criar a RNC"
If txtof = "" Then
    NomeCampo = "a ordem"
    ProcVerificaAcao
    txtof.SetFocus
    Exit Sub
End If
If cmbOS = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    cmbOS.SetFocus
    Exit Sub
End If
'If Novo_CQNC = True Then
'    usMsgbox ("Salve a não conformidade antes de criar a RNC."), vbExclamation, "CAPRIND v5.0"
'    cmdSalvar.SetFocus
'    Exit Sub
'End If
If txtRNC = "" Then
    USMsgBox ("Crie a RNC antes de visualizar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
RNC_Controle_Medicao = False
RNC_Inspecao_Recebimento = False
RNC_Nao_Conformidade = True
RNC_Solicitacao_Desvio = False
Sit_REG = 1
frmQualidade_RNC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcReposicaoRetrabalho()
On Error GoTo tratar_erro

If txtidos = "" Then
    USMsgBox ("Informe a não conformidade antes de visualizar a(s) ordem(ns) de reposição e a(s) Os(s) de retrabalho."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmcqnc_ListaRetrabalho.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdSD_Click()
On Error GoTo tratar_erro

Acao = "criar a SD"
If txtof = "" Then
    NomeCampo = "a ordem"
    ProcVerificaAcao
    txtof.SetFocus
    Exit Sub
End If
If cmbOS = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    cmbOS.SetFocus
    Exit Sub
End If
If Novo_CQNC = True Then
    USMsgBox ("Salve a não conformidade antes de criar a SD."), vbExclamation, "CAPRIND v5.0"
    CmdSalvar.SetFocus
    Exit Sub
End If
RNC_Nao_Conformidade = True
frmCQ_SD.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcGravar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF8: ProcReposicaoRetrabalho
    Case vbKeyF9: procOsRetrabalho
    Case vbKeyF10: ProcRNC
    Case vbKeyF11: procDisposicao
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 14, True
If PCP_Ordem = True Then
    Caption = "PCP - Não conformidade"
    Formulario = "PCP/Não conformidade"
    With USToolBar1
        USToolBar1.ButtonState(1) = 5
        USToolBar1.ButtonState(3) = 5
        USToolBar1.ButtonState(4) = 5
        USToolBar1.ButtonState(7) = 0
        USToolBar1.ButtonState(10) = 5
    End With
    Cmb_opcao_lista.AddItem "Ordem de retrabalho"
Else
    Cmb_opcao_lista.AddItem "Excluir"
    Formulario = "Qualidade/Não conformidade"
    USToolBar1.ButtonState(7) = 5
End If
Direitos
ProcLimpaVariaveisPrincipais
ProcLimpaCampos

If PCP_Ordem = True Then
    Cmb_opcao_lista = "Ordem de retrabalho"
Else
    Cmb_opcao_lista = "Excluir"
End If

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Não conformidade"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaComboOperador()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from usuarios where bloqueado = 'False' order by usuario", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    cmbOperador.AddItem ""
    Do While TBAbrir.EOF = False
        If IsNull(TBAbrir!CODIGO) = False And TBAbrir!CODIGO <> "" Then OperadorTexto = TBAbrir!CODIGO & "-" & TBAbrir!Usuario Else OperadorTexto = TBAbrir!Usuario
        cmbOperador.AddItem OperadorTexto
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaComboMaquina()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CadMaquinas order by Maquina", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Cmb_maquina.AddItem TBAbrir!maquina
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaComboOrigem(Combo As ComboBox)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CQ_NC_FABRICA_origem order by Origem", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .AddItem ""
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Origem
            .ItemData(.NewIndex) = TBAbrir!ID
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaComboCausa(Combo As ComboBox)
On Error GoTo tratar_erro

With Combo
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CQ_NC_FABRICA_causa order by Causa", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .AddItem ""
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!Causa
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_CQNC = True Then
    If USMsgBox("A não conformidade ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar
        If Novo_CQNC = True Then Exit Sub Else Unload Me
    End If
End If
Novo_CQNC = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtidos = ""
txtof.Text = ""
cmbOS.Clear
Cmb_maquina.Clear
txtFase.Text = ""
txtData.Value = Date
txthora.Value = Format(Now, "hh:mm:ss")
Txt_turno = 0
txtpubusuario.Text = pubUsuario
txtdesenho.Text = ""
txtreferencia = ""
txtdescricao.Text = ""
cmbOperador.Clear
txtLote.Text = ""
txtQuant.Text = ""
txtaprovadas.Text = ""
txtnc.Text = ""
Txt_ID_RNC = 0
txtRNC.Text = ""
txtID_SD = ""
chkAnalizada.Value = 0
Cmb_origem.Clear
txtParecerF = ""
Opt_aprovado.Value = False
Opt_aprovado_desvio.Value = False
Opt_rejeitar.Value = False
Opt_retrabalhar.Value = False
Opt_selecionar.Value = False
Opt_reaproveitar.Value = False
optDevolver.Value = False
Opt_outros.Value = False
Opt_nada_consta.Value = True
txtobscq.Text = ""
ProcCarregaComboMaquina
ProcCarregaComboOperador
ProcCarregaComboOrigem Cmb_origem
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

MesmaOrdem = 0
MesmaOS = 0
MesmoItem = ""
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaFases.ListItems.Clear
If StrSql_CQ_NC = "" Then Exit Sub
'Debug.print StrSql_CQ_NC

Set TBLISTA_NC = CreateObject("adodb.recordset")
TBLISTA_NC.Open StrSql_CQ_NC, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_NC.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaFases.ListItems.Clear
TBLISTA_NC.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_NC.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_NC.PageSize
ContadorReg = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_NC.RecordCount - IIf(Pagina > 1, (TBLISTA_NC.PageSize * (Pagina - 1)), 0), TBLISTA_NC.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_NC.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaFases.ListItems.Add(, , TBLISTA_NC!CODIGO)
        .SubItems(1) = IIf(IsNull(TBLISTA_NC!IDProducao), "", TBLISTA_NC!IDProducao)
        .SubItems(2) = IIf(IsNull(TBLISTA_NC!Ordem), "", TBLISTA_NC!Ordem)
        .SubItems(3) = IIf(IsNull(TBLISTA_NC!OS), "", TBLISTA_NC!OS)
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select desenho from producao where ordem = " & TBLISTA_NC!Ordem, Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            .SubItems(4) = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
        End If
        TBOrdem.Close
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select Fase from ordemservico where idproducao = " & TBLISTA_NC!OS, Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            .SubItems(5) = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
        End If
        TBOrdem.Close
        .SubItems(6) = IIf(IsNull(TBLISTA_NC!Quant), "0,00", Format(TBLISTA_NC!Quant, "###,##0.00"))
        .SubItems(7) = IIf(IsNull(TBLISTA_NC!TTNC), "0,00", Format(TBLISTA_NC!TTNC, "###,##0.00"))
        .SubItems(8) = IIf(IsNull(TBLISTA_NC!QTCD), "0,00", Format(TBLISTA_NC!QTCD, "###,##0.00"))
        .SubItems(9) = IIf(IsNull(TBLISTA_NC!Data), "", Format(TBLISTA_NC!Data, "dd/mm/yy"))
        .SubItems(10) = IIf(IsNull(TBLISTA_NC!Operador), "", TBLISTA_NC!Operador)
        .SubItems(11) = IIf(IsNull(TBLISTA_NC!ParecerCQ), "", TBLISTA_NC!ParecerCQ)
        If TBLISTA_NC!Analizada = False Then
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
        End If
    End With
    TBLISTA_NC.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_NC.RecordCount
If TBLISTA_NC.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_NC.PageCount
ElseIf TBLISTA_NC.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_NC.PageCount & " de: " & TBLISTA_NC.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_NC.AbsolutePage - 1 & " de: " & TBLISTA_NC.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub



Private Sub ListaFases_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

MesmaOrdem = 0
MesmaOS = 0
MesmoItem = ""
If ColumnHeader = "" Then
    With ListaFases
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select OS.IDProducao FROM Ordemservico OS INNER JOIN Producaofases PF ON OS.idproducao = PF.OS where OS.Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and OS.fase = " & .ListItems.Item(InitFor).ListSubItems(5) & " and OS.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then
                        TBOrdem.Close
                        GoTo Proximo
                    End If
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select OS.IDProducao FROM Ordemservico OS INNER JOIN Producaofases_Backup PFB ON OS.idproducao = PFB.OS where OS.Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and OS.fase = " & .ListItems.Item(InitFor).ListSubItems(5) & " and OS.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then
                        TBOrdem.Close
                        GoTo Proximo
                    End If
                    TBOrdem.Close
                    If .ListItems.Item(InitFor).ListSubItems(1) <> "" And .ListItems.Item(InitFor).ListSubItems(1) <> "0" Then
                        Set TBOrdem = CreateObject("adodb.recordset")
                        TBOrdem.Open "Select Descricao, TempoInicio from Producaofases where IDProducao = " & .ListItems.Item(InitFor).ListSubItems(1) & " and OS = " & .ListItems.Item(InitFor).ListSubItems(3), Conexao, adOpenKeyset, adLockOptimistic
                        If TBOrdem.EOF = False Then
                            GoTo Proximo
                        Else
                            Set TBOrdem = CreateObject("adodb.recordset")
                            TBOrdem.Open "Select Descricao, TempoInicio from Producaofases_Backup where IDProducao = " & .ListItems.Item(InitFor).ListSubItems(1) & " and OS = " & .ListItems.Item(InitFor).ListSubItems(3), Conexao, adOpenKeyset, adLockOptimistic
                            If TBOrdem.EOF = False Then
                                GoTo Proximo
                            End If
                        End If
                        TBOrdem.Close
                    End If
                ElseIf Cmb_opcao_lista = "RNC" Then
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select ID_RNC from CQ_NC_FABRICA where Codigo = " & .ListItems(InitFor) & " and ID_RNC is not null and ID_RNC <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then GoTo Proximo
                    TBOrdem.Close
                    
                    If MesmaOrdem = 0 Then MesmaOrdem = .ListItems.Item(InitFor).ListSubItems(2)
                    If MesmaOrdem <> .ListItems.Item(InitFor).ListSubItems(2) Then GoTo Proximo
                ElseIf Cmb_opcao_lista = "Ordem de retrabalho" Then
                    'Retrabalho
                    If .ListItems.Item(InitFor).ListSubItems(10) <> "Retrabalhar" Then GoTo Proximo
                
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select OrdemRetrabalho from CQ_NC_FABRICA where Codigo = " & .ListItems.Item(InitFor) & " and OrdemRetrabalho IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
                    If TBOrdem.EOF = False Then
                        TBOrdem.Close
                        GoTo Proximo
                    End If
                    TBOrdem.Close
                    
                    Set TBOSC = CreateObject("adodb.recordset")
                    TBOSC.Open "Select Ordem from Ordemservico where Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and fase = " & .ListItems.Item(InitFor).ListSubItems(5) & " and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockReadOnly
                    If TBOSC.EOF = False Then
                        TBOSC.Close
                        GoTo Proximo
                    End If
                    TBOSC.Close
                    
                    If MesmoItem = "" Then MesmoItem = .ListItems.Item(InitFor).ListSubItems(4)
                    If MesmoItem <> .ListItems.Item(InitFor).ListSubItems(4) Then GoTo Proximo
                ElseIf Cmb_opcao_lista = "Disposição" Then
                    If MesmaOS = 0 Then MesmaOS = .ListItems.Item(InitFor).ListSubItems(3)
                    If MesmaOS <> .ListItems.Item(InitFor).ListSubItems(3) Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaFases, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Listafases_DblClick()
On Error GoTo tratar_erro

btnNS_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Listafases_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Permitido = False
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Cmb_opcao_lista = "Excluir" Then
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select OS.IDProducao FROM Ordemservico OS INNER JOIN Producaofases PF ON OS.idproducao = PF.OS where OS.Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and OS.fase = " & .ListItems.Item(InitFor).ListSubItems(5) & " and OS.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    USMsgBox ("Não é permitido excluir esta não conformidade, pois a mesma está sendo utilizada no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBOrdem.Close
                    Exit Sub
                End If
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select OS.IDProducao FROM Ordemservico OS INNER JOIN Producaofases_Backup PFB ON OS.idproducao = PFB.OS where OS.Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and OS.fase = " & .ListItems.Item(InitFor).ListSubItems(5) & " and OS.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    USMsgBox ("Não é permitido excluir esta não conformidade, pois a mesma está sendo utilizada no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
                    TBOrdem.Close
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                TBOrdem.Close
                If .ListItems.Item(InitFor).ListSubItems(1) <> "" And .ListItems.Item(InitFor).ListSubItems(1) <> "0" Then
                    Texto = ""
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select Descricao, TempoInicio from Producaofases where IDProducao = " & .ListItems.Item(InitFor).ListSubItems(1) & " and OS = " & .ListItems.Item(InitFor).ListSubItems(3), Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = False Then
                        Texto = TBOrdem!Descricao & " - " & TBOrdem!TempoInicio
                    Else
                        Set TBOrdem = CreateObject("adodb.recordset")
                        TBOrdem.Open "Select Descricao, TempoInicio from Producaofases_Backup where IDProducao = " & .ListItems.Item(InitFor).ListSubItems(1) & " and OS = " & .ListItems.Item(InitFor).ListSubItems(3), Conexao, adOpenKeyset, adLockOptimistic
                        If TBOrdem.EOF = False Then
                            Texto = TBOrdem!Descricao & " - " & TBOrdem!TempoInicio
                        End If
                    End If
                    TBOrdem.Close
                    If Texto <> "" Then
                        USMsgBox ("Não é permitido excluir esta não conformidade, pois a mesma está vinculada ao apontamento " & Texto & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                End If
            ElseIf Cmb_opcao_lista = "RNC" Then
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select ID_RNC from CQ_NC_FABRICA where Codigo = " & .ListItems(InitFor) & " and ID_RNC is not null and ID_RNC <> 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    USMsgBox ("Não é permitido criar RNC para esta não conformidade, pois a mesma já possui RNC."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                TBOrdem.Close
                
                If MesmaOrdem = 0 Then MesmaOrdem = .ListItems.Item(InitFor).ListSubItems(2)
                If MesmaOrdem <> .ListItems.Item(InitFor).ListSubItems(2) Then
                    USMsgBox ("Não é possível criar RNC para ordens diferentes."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            ElseIf Cmb_opcao_lista = "Disposição" Then
                If MesmaOS = 0 Then MesmaOS = .ListItems.Item(InitFor).ListSubItems(3)
                If MesmaOS <> .ListItems.Item(InitFor).ListSubItems(3) Then
                    USMsgBox ("Não é possível alterar disposição de OS's diferentes."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            ElseIf Cmb_opcao_lista = "Ordem de retrabalho" Then
                'Retrabalho
                If .ListItems.Item(InitFor).ListSubItems(11) <> "Retrabalhar" Then
                    USMsgBox "Não é permitido criar ordem de retrabalho, pois a disposição é diferente de retrabalhar.", vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select OrdemRetrabalho from CQ_NC_FABRICA where Codigo = " & .ListItems.Item(InitFor) & " and OrdemRetrabalho IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
                If TBOrdem.EOF = False Then
                    USMsgBox "Não é permitido criar ordem de retrabalho, pois já existe ordem de retrabalho criada para essa não conformidade.", vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBOrdem.Close
                    Exit Sub
                End If
                TBOrdem.Close
                
                Set TBOSC = CreateObject("adodb.recordset")
                TBOSC.Open "Select Ordem from Ordemservico where Ordem = " & .ListItems.Item(InitFor).ListSubItems(2) & " and fase = " & .ListItems.Item(InitFor).ListSubItems(5) & " and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockReadOnly
                If TBOSC.EOF = False Then
                    USMsgBox "Não é permitido criar ordem de retrabalho, pois já existe OS de retrabalho criada para essa não conformidade.", vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    TBOSC.Close
                    Exit Sub
                End If
                TBOSC.Close
                
                If MesmoItem = "" Then MesmoItem = .ListItems.Item(InitFor).ListSubItems(4)
                If MesmoItem <> .ListItems.Item(InitFor).ListSubItems(4) Then
                    USMsgBox ("Não é possível criar ordem de retrabalho para não conformidades de produtos diferentes."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            Permitido = True
        End If
    Next InitFor
End With
If Permitido = False Then
    MesmaOrdem = 0
    MesmaOS = 0
    MesmoItem = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaFases_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaFases.ListItems.Count = 0 Then Exit Sub
txtId = ListaFases.SelectedItem
IDProducao = ListaFases.SelectedItem.ListSubItems.Item(1).Text

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from cq_nc_fabrica where Codigo = " & ListaFases.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    If PCP_Ordem = False Then
        Frame1.Enabled = True
        Frame3.Enabled = True
    End If
    Novo_CQNC = False
    CodigoLista = ListaFases.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
'TBOrdem.Open "Select P.QUANT,P.QTCD, P.Ordem, P.Desenho, P.N_referencia, P.Produto, OS.Fase, OS.QTOK from ordemservico OS INNER JOIN producao P ON P.Ordem = OS.Ordem where OS.idproducao = " & TBLISTA!OS, Conexao, adOpenKeyset, adLockOptimistic
TBOrdem.Open "Select P.Quant , P.QUANTProd ,P.QTCD, P.Ordem, P.Desenho, P.N_referencia, P.Produto, OS.Fase, OS.QTOK from ordemservico OS INNER JOIN producao P ON P.Ordem = OS.Ordem where OS.idproducao = " & TBLISTA!OS, Conexao, adOpenKeyset, adLockOptimistic

If TBOrdem.EOF = False Then
    txtof.Text = IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
    ProcCarregaComboOS
    txtFase.Text = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
    txtaprovadas.Text = IIf(IsNull(TBOrdem!QTOK), 0, Format(TBOrdem!QTOK, "###,##0.0000"))
    txtQTCD.Text = IIf(IsNull(TBOrdem!QTCD), 0, Format(TBOrdem!QTCD, "###,##0.0000"))
    txtdescricao.Text = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
    txtdesenho.Text = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
    txtreferencia = IIf(IsNull(TBOrdem!N_referencia), "", TBOrdem!N_referencia)
End If
txtidos.Text = TBLISTA!CODIGO
If IsNull(TBLISTA!OS) = False And TBLISTA!OS <> "" Then cmbOS = TBLISTA!OS
txtData.Value = TBLISTA!Data
txthora.Value = Format(TBLISTA!Hora, "hh:mm:ss")
Txt_turno = IIf(IsNull(TBLISTA!Turno), 0, TBLISTA!Turno)
txtpubusuario.Text = IIf(IsNull(TBLISTA!Usuario), pubUsuario, TBLISTA!Usuario)
TxtdescricaoNC = IIf(IsNull(TBLISTA!obsFab), "", TBLISTA!obsFab)
1:
    NomeCampo = "o operador"
    If IsNull(TBLISTA!Operador) = False And TBLISTA!Operador <> "" Then cmbOperador = TBLISTA!Operador
    NomeCampo = "o posto de trabalho"
    If IsNull(TBLISTA!maquina) = False And TBLISTA!maquina <> "" Then Cmb_maquina = TBLISTA!maquina
    
2:
    If IsNull(TBLISTA!ID_origem) = False And TBLISTA!ID_origem <> "" Then
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select Origem from CQ_NC_FABRICA_origem where ID = " & TBLISTA!ID_origem, Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            Cmb_origem = TBproducao!Origem
        End If
        TBproducao.Close
    End If
        
    txtLote.Text = Format(TBOrdem!QuantProd, "###,##0.0000")
    txtQuant.Text = Format(TBOrdem!Quant, "###,##0.0000")
    
    txtnc.Text = Format(TBLISTA!TTNC, "###,##0.0000")
    txtParecerF.Text = IIf(IsNull(TBLISTA!PARECERFAB), "", TBLISTA!PARECERFAB)
    If IsNull(TBLISTA!ParecerCQ) = False And TBLISTA!ParecerCQ <> "" Then
        Select Case TBLISTA!ParecerCQ
            Case "Aprovado": Opt_aprovado.Value = True
            Case "Aprovado c/ desvio": Opt_aprovado_desvio.Value = True
            Case "Rejeitar": Opt_rejeitar.Value = True
            Case "Retrabalhar": Opt_retrabalhar.Value = True
            Case "Selecionar": Opt_selecionar.Value = True
            Case "Reaproveitar": Opt_reaproveitar.Value = True
            Case "Devolver": optDevolver.Value = True
            Case "Outros": Opt_outros.Value = True
            Case "Nada consta": Opt_nada_consta.Value = True
        End Select
    End If
    txtobscq.Text = IIf(IsNull(TBLISTA!obsCQ), "", TBLISTA!obsCQ)
    
    Txt_ID_RNC = IIf(IsNull(TBLISTA!ID_RNC), 0, TBLISTA!ID_RNC)
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select ID_texto, Seq FROM CQ_RNC where ID = " & Txt_ID_RNC, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        txtRNC = IIf(IsNull(TBCompras_Pedido!Seq), TBCompras_Pedido!id_texto, TBCompras_Pedido!id_texto & "/" & IIf(TBCompras_Pedido!Seq < 10, "0" & TBCompras_Pedido!Seq, TBCompras_Pedido!Seq))
    End If
    TBCompras_Pedido.Close
    
    txtID_SD = IIf(IsNull(TBLISTA!ID_SD), "", TBLISTA!ID_SD)
        
    If TBLISTA!Analizada = True Then chkAnalizada.Value = 1 Else chkAnalizada.Value = 0
    If TBLISTA!IDProducao = 0 Then
        With cmbOperador
            .Locked = False
            .TabStop = True
        End With
        With txtnc
            .Locked = False
            .TabStop = True
        End With
    Else
        With cmbOperador
            .Locked = True
            .TabStop = False
        End With
        With txtnc
            .Locked = True
            .TabStop = False
        End With
    End If
    Novo_CQNC = False
TBOrdem.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        If NomeCampo = "o operador" Then
            With cmbOperador
                .Clear
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from usuarios where bloqueado = 'False' order by usuario", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .AddItem ""
                    Do While TBAbrir.EOF = False
                        .AddItem TBAbrir!CODIGO & "-" & TBAbrir!Usuario
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
                .AddItem TBLISTA!Operador
            End With
            GoTo 1
        Else
            USMsgBox ("Não foi encontrado " & NomeCampo & " dessa não conformidade."), vbExclamation, "CAPRIND v5.0"
            GoTo 2
        End If
        
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcGravar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtof.Text = "" Then
    NomeCampo = "a ordem"
    ProcVerificaAcao
    txtof.SetFocus
    Exit Sub
End If
If cmbOS = "" Then
    NomeCampo = "a OS"
    ProcVerificaAcao
    cmbOS.SetFocus
    Exit Sub
End If
If Cmb_maquina = "" Then
    NomeCampo = "o posto de trabalho"
    ProcVerificaAcao
    Cmb_maquina.SetFocus
    Exit Sub
End If
If txthora.Value = "00:00:00" Then
    NomeCampo = "a hora"
    ProcVerificaAcao
    txthora.SetFocus
    Exit Sub
End If
If cmbOperador = "" Then
    NomeCampo = "o operador"
    ProcVerificaAcao
    cmbOperador.SetFocus
    Exit Sub
End If
If txtaprovadas = "" Then
    NomeCampo = "a quantidade aprovada"
    ProcVerificaAcao
    txtaprovadas.SetFocus
    Exit Sub
End If
valor = IIf(txtnc = "", 0, txtnc)
If valor <= 0 Then
    NomeCampo = "a quantidade não conforme"
    ProcVerificaAcao
    txtnc.SetFocus
    Exit Sub
End If
If Opt_aprovado.Value = False And Opt_aprovado_desvio.Value = False And Opt_rejeitar.Value = False And Opt_retrabalhar.Value = False And Opt_selecionar.Value = False And Opt_reaproveitar.Value = False And optDevolver.Value = False And Opt_outros.Value = False And Opt_nada_consta.Value = False Then
    NomeCampo = "a disposição"
    ProcVerificaAcao
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from cq_nc_fabrica where Codigo = " & IIf(txtidos = "", 0, txtidos), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    StrSql_CQ_NC = "Select * from cq_nc_fabrica where analizada = 'False' order by os"
Else
    If TBGravar!ParecerCQ <> "Nada consta" And Opt_nada_consta.Value = True Then
        If TBGravar!ParecerCQ = "Retrabalhar" Or TBGravar!ParecerCQ = "Selecionar" Then
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select * FROM Ordemservico INNER JOIN Producaofases ON Ordemservico.idproducao = Producaofases.OS where Ordemservico.Idproducao = " & cmbOS & " and Ordemservico.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = False Then
                USMsgBox ("Não é permitido alterar esta não conformidade, pois a mesma está sendo utilizada no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
                TBOrdem.Close
                Exit Sub
            End If
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select * FROM Ordemservico INNER JOIN Producaofases_Backup ON Ordemservico.idproducao = Producaofases_Backup.OS where Ordemservico.Idproducao = " & cmbOS & " and Ordemservico.retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = False Then
                USMsgBox ("Não é permitido alterar esta não conformidade, pois a mesma está sendo utilizada no módulo PCP/Gerenciamento de ordem."), vbExclamation, "CAPRIND v5.0"
                TBOrdem.Close
                Exit Sub
            End If
            TBOrdem.Close
        End If
    End If
End If

If Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True Or Opt_rejeitar.Value = True Or Opt_retrabalhar.Value = True Or Opt_selecionar.Value = True Or Opt_reaproveitar.Value = True Or optDevolver.Value = True Or Opt_outros.Value = True Then
    If USMsgBox("Essa não conformidade já foi analisada?", vbYesNo, "CAPRIND v5.0") = vbYes Then chkAnalizada.Value = 1 Else chkAnalizada.Value = 0
End If

TBGravar!Ordem = txtof
TBGravar!OS = cmbOS
TBGravar!maquina = Cmb_maquina
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from CadMaquinas where Maquina = '" & Cmb_maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    TBGravar!Setor = TBMaquinas!Setor
End If
TBMaquinas.Close
TBGravar!Data = txtData.Value
TBGravar!Hora = txthora.Value
TBGravar!Turno = Txt_turno
TBGravar!Usuario = txtpubusuario
TBGravar!Operador = cmbOperador
TBGravar!LOTE = txtLote.Text
TBGravar!TTNC = txtnc.Text
TBGravar!ID_RNC = Txt_ID_RNC
TBGravar!ID_SD = IIf(txtID_SD = "", 0, txtID_SD)
If txtParecerF.Text = "" Then TBGravar!PARECERFAB = Null Else TBGravar!PARECERFAB = txtParecerF.Text
If Cmb_origem <> "" Then TBGravar!ID_origem = Cmb_origem.ItemData(Cmb_origem.ListIndex)
TBGravar!obsFab = IIf(TxtdescricaoNC = "", Null, TxtdescricaoNC)

ProcGravarNCOSMaq

If Opt_nada_consta.Value = True Then
    TBGravar!ParecerCQ = "Nada consta"
    chkAnalizada.Value = 0
End If
If Opt_aprovado.Value = True Then TBGravar!ParecerCQ = "Aprovado"
If Opt_aprovado_desvio.Value = True Then TBGravar!ParecerCQ = "Aprovado c/ desvio"
If Opt_rejeitar.Value = True Then TBGravar!ParecerCQ = "Rejeitar"
If Opt_retrabalhar.Value = True Then TBGravar!ParecerCQ = "Retrabalhar"
If Opt_selecionar.Value = True Then TBGravar!ParecerCQ = "Selecionar"
If Opt_reaproveitar.Value = True Then TBGravar!ParecerCQ = "Reaproveitar"
If optDevolver.Value = True Then TBGravar!ParecerCQ = "Devolver"
If Opt_outros.Value = True Then TBGravar!ParecerCQ = "Outros"

If chkAnalizada.Value = 1 Then TBGravar!Analizada = True Else TBGravar!Analizada = False

TBGravar!obsCQ = txtobscq.Text
TBGravar.Update
txtidos = TBGravar!CODIGO

ProcGravarNCOrdem txtof

TBGravar.Close

If Novo_CQNC = True Then
    USMsgBox ("Novo registro cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_CQ_NC = "Select * from cq_nc_fabrica where Codigo = " & txtidos
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And ListaFases.ListItems.Count <> 0 Then
        ListaFases.SelectedItem = ListaFases.ListItems(CodigoLista)
        ListaFases.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Qualidade/Não conformidade"
    ID_documento = txtidos
    Documento = "Data: " & txtData.Value & " - Hora: " & txthora.Value & " - Ordem: " & txtof & " - OS: " & cmbOS & " - Cód. interno: " & txtdesenho & " - Operador: " & cmbOperador
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_CQNC = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCriarOSRetrabalho()
On Error GoTo tratar_erro

frmcqnc_retrabalho.Show 1
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select OS.*, P.Desenho from Ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.Ordem = " & txtof & " and OS.fase = " & txtFase, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    'Busca dados das fases do processo
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * from fases where idfase = " & TBOrdem!IDFase, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        Set TBOSC = CreateObject("adodb.recordset")
        TBOSC.Open "Select * from Ordemservico where Ordem = " & txtof & " and fase = " & txtFase & " and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBOSC.EOF = True Then TBOSC.AddNew
        TBOSC!Ordem = txtof
        TBOSC!IDFase = TBFases!IDFase
        TBOSC!quantidade = QuantSolicitado
        TBOSC!Fase = TBFases!Fase
        TBOSC!maquina = TBOrdem!maquina
        DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * QuantSolicitado) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
        TBOSC!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
        TBOSC!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
        TBOSC!Pronto = "NÃO"
        TBOSC!Preparacao = IIf(IsNull(TBFases!Preparacao), "00:00:00", TBFases!Preparacao)
        TBOSC!Execucao = IIf(IsNull(TBFases!Execucao), "00:00:00", TBFases!Execucao)
        TBOSC!IDPROCESSO = TBFases!IDPROCESSO
        TBOSC!PrazoFinal = TBOrdem!PrazoFinal
        TBOSC!descfase = FamiliaAntiga
        TBOSC!Obs = Familiatext
        TBOSC!TempoPreparacao = TBFases!TempoPreparacao
        TBOSC!TempoExecucao = TBFases!TempoExecucao
        TBOSC!OSControlada = TBOrdem!OSControlada
        TBOSC!Processo_controlado = TBOrdem!Processo_controlado
        TBOSC!custos = TBOrdem!custos
        
        If TBFases!pecahora = True Then
            TBOSC!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        Else
            If IsNull(TBOSC!Execucao) = False And TBOSC!Execucao <> "00:00:00" Then
                ElapsedTime (TBOSC!Execucao)
                TBOSC!Pcshora = 3600 / s
            End If
        End If
        TBOSC!pc_te = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        
        TBOSC!status = "Aguardando"
        TBOSC!Retrabalho = True
        
        'Verifica custo previsto da os
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & TBOrdem!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            TotalFaseSeg = 0
            CustoFase = 0
            CustohoraSeg = 0
            TotalPreparacaoSeg = 0
            CustopreparacaoSeg = 0
            
            'Verifica custo de execucao por segundos * custo da hora maquina
            CustohoraSeg = TBMaquinas!PrecoHora / 3600
            ElapsedTime (TBOrdem!Execucao)
            TotalFaseSeg = s
            CustoFase = CustohoraSeg * TotalFaseSeg
            
            'Verifica custo de preparacao por segundos * custo da hora maquina
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then CustohoraSeg = TBMaquinas!PrecoHora_Setup / 3600
            ElapsedTime (TBOrdem!Preparacao)
            TotalPreparacaoSeg = s
            CustopreparacaoSeg = CustohoraSeg * TotalPreparacaoSeg
            
            TBOSC!CPPECA = Format(CustoFase + (CustopreparacaoSeg / QuantSolicitado), "###,##0.0000000000")
            TBOSC!CPLOTE = Format(TBOSC!CPPECA * QuantSolicitado, "###,##0.00")
        End If
        TBMaquinas.Close
        
        'Verif. se tem plano de inspeção
        TBOSC!IDPlano = 0
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select * from Plano where Desenho = '" & TBOrdem!Desenho & "' and Fase = " & TBOSC!Fase, Conexao, adOpenKeyset, adLockOptimistic
        If TBplano.EOF = False Then
            TBOSC!IDPlano = TBplano!IDPlano
        End If
        TBplano.Close
        TBOSC.Update
        TBOSC.Close
    End If
    TBFases.Close
End If
TBOrdem.Close

'Atualiza custo previsto da ordem
CustoOrdem = 0
TotalOrdem = 0
PcHora = 0
Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select * from ordemServico where Ordem = " & txtof & " ORDER BY FASE", Conexao, adOpenKeyset, adLockOptimistic
If TBOSC.EOF = False Then
    TBOSC.MoveFirst
    Do Until TBOSC.EOF = False
        TOTALPECA = 0
        TotalOS = 0
        CustoOS = 0
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select * FROM FASES WHERE IDFASE = " & TBOSC!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            PcHora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
            
            'Tempo total por peça
            ElapsedTime (IIf(IsNull(TBFases!Execucao), 0, TBFases!Execucao))
            If PcHora <> 0 Then TOTALPECA = TOTALPECA + (s / PcHora)
            
            'Tempo total do lote
            If PcHora <> 0 Then TotalOS = s / PcHora Else TotalOS = 0
            ElapsedTime (IIf(IsNull(TBFases!Preparacao), 0, TBFases!Preparacao))
            TotalOS = (TotalOS * TBOSC!quantidade) + s
            
            'Custo total do lote
            CustoOS = (TBFases!Custo * TBOSC!quantidade) + IIf(IsNull(TBFases!Custoprep), 0, TBFases!Custoprep)
            CustoOrdem = CustoOrdem + CustoOS
            
            TBOSC!pecahora = TBFases!pecahora
            If TBFases!pecahora = True Then
                TBOSC!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
            Else
                If IsNull(TBOSC!Execucao) = False And TBOSC!Execucao <> "00:00:00" Then
                    ElapsedTime (TBOSC!Execucao)
                    TBOSC!Pcshora = 3600 / s
                End If
            End If
            'Tempo total por peça
            TBOSC!TempoExecucao = TOTALPECA
            TBOSC!TempoExecucao = FormataTempo(TBOSC!TempoExecucao)
            
            TBOSC!TTLPREVS = TotalOS 'Tempo total do lote previsto em segundos
            TBOSC!TempoTotalLote = FormataTempo(TBOSC!TTLPREVS) 'Tempo total do lote previsto
            
            'Custo por peça
            If TBOSC!quantidade <> 0 Then TBOSC!CPPECA = Format(TBFases!Custo + (IIf(IsNull(TBFases!Custoprep), 0, TBFases!Custoprep) / TBOSC!quantidade), "###,##0.0000000000") Else TBOSC!CPPECA = Format(TBFases!Custo + IIf(IsNull(TBFases!Custoprep), 0, TBFases!Custoprep), "###,##0.0000000000")
            
            'Custo do lote
            TBOSC!CPLOTE = Format(CustoOS, "###,##0.00")
            
            TotalOrdem = TotalOrdem + TotalOS
            TBOSC.Update
        End If
        TBFases.Close
        TBOSC.MoveNext
    Loop
End If
TBOSC.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from producao where Ordem = " & txtof, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    QuantSolicitado = TBAbrir!Quant
    
    'Custo por peça
    If Int(QuantSolicitado) <> 0 Then TBAbrir!cpp = CustoOrdem / Int(QuantSolicitado) Else TBAbrir!cpp = CustoOrdem
    'Custo do lote
    TBAbrir!CTTPrev = CustoOrdem
    'Tempo total por peça
    If TotalOrdem <> 0 Then
        TBAbrir!TPP = TotalOrdem / Int(QuantSolicitado)
        TBAbrir!TPP = FormataTempo(TBAbrir!TPP)
    Else
        TBAbrir!TPP = "00:00:00"
    End If
    'Tempo total do lote
    TBAbrir!TTTPrev = TotalOrdem
    TBAbrir!TTTPrev = FormataTempo(TBAbrir!TTTPrev)
    'Tempo total do lote em segundos
    TBAbrir!TTTPREVSegundos = TotalOrdem
    TBAbrir.Update
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCriarOrdemRetrabalho(OrdemCriar As Long)
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from Producao where Ordem = " & OrdemCriar, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from producao", Conexao, adOpenKeyset, adLockOptimistic
    TBproducao.AddNew
    TBproducao!Ordem = funcCriarNumeroOrdem
    TBproducao!status = "Aberta"
    TBproducao!Data_cadastro = Date
    TBproducao!Responsavel = pubUsuario
    TBproducao!Impof = 0
    If Permitido1 = True Then
        TBproducao!Cliente = TBOrdem!Cliente
        TBproducao!IDCliente = TBOrdem!IDCliente
    End If
    TBproducao!PrazoEntrega = PrazoEntregaOrdem
    TBproducao!Quant = QtdeOrdemRetrabalho
    TBproducao!Saldo = True
    TBproducao!ID_empresa = TBOrdem!ID_empresa
    TBproducao!IDPROCESSO = TBOrdem!IDPROCESSO
    TBproducao!Data = Date
    TBproducao!Desenho = TBOrdem!Desenho
    TBproducao!Revitem = TBOrdem!Revitem
    TBproducao!Produto = TBOrdem!Produto
    TBproducao!Codigo_produto = TBOrdem!Codigo_produto
    TBproducao!Rev_produto = TBOrdem!Rev_produto
    TBproducao!N_referencia = TBOrdem!N_referencia
    TBproducao!pronta = "NÃO"
    TBproducao!Tipo = TBOrdem!Tipo
    TBproducao!IMPREQ = TBOrdem!IMPREQ
    TBproducao!Consignacao = TBOrdem!Consignacao
    TBproducao!OSControlada = TBOrdem!OSControlada
    TBproducao!Processo_controlado = TBOrdem!Processo_controlado
    TBproducao!Tipo_Processo = TBOrdem!Tipo_Processo
    TBproducao!Entrar_estoque = TBOrdem!Entrar_estoque
    TBproducao!Retirar_estoque = TBOrdem!Retirar_estoque
    TBproducao!Escopo = TBOrdem!Escopo
    TBproducao!Obs = "Ordem de retrabalho"
    TBproducao!Retrabalho = True
    TBproducao.Update
    OrdemRetrabalho = TBproducao!Ordem
    TBproducao.Close
        
    ProcCriarOrdemServico TBOrdem!IDPROCESSO
End If
TBOrdem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCriarOrdemServico(IDprocessoCriar As Long)
On Error GoTo tratar_erro

TotalUtilizado = "00:00:00"
'Busca dados das fases do processo
Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select F.* from fases F INNER JOIN Processos P ON P.IDProcesso = F.IDProcesso where F.idprocesso = " & IDprocessoCriar & " AND F.versao = 'A' and P.DtValidacao IS NOT NULL order by F.fase", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    Do While TBFases.EOF = False
        Set TBProducaoFases = CreateObject("adodb.recordset")
        TBProducaoFases.Open "Select * from ordemservico where Ordem = " & OrdemRetrabalho & " and IdFase = " & TBFases!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBProducaoFases.EOF = True Then
            TBProducaoFases.AddNew
            If TBFases!Nao_aponta = True Then
                TBProducaoFases!Pronto = "SIM"
                TBProducaoFases!status = "Concluída"
                TBProducaoFases!DataConclusao = Date
            Else
                TBProducaoFases!Pronto = "NÃO"
                TBProducaoFases!status = "Aguardando"
            End If
        End If
        TBProducaoFases!Fase = TBFases!Fase
        TBProducaoFases!Rev_Fase = IIf(IsNull(TBFases!Revisao), 0, TBFases!Revisao)
        TBProducaoFases!Grupo_op = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)
        TBProducaoFases!IDFase = TBFases!IDFase
        TBProducaoFases!IDPlano = FunVerifIDPlano(TBFases!IDFase)
        TBProducaoFases!Retrabalho = True
        
        TBProducaoFases!maquina = TBFases!maquina
        
        DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * QtdeOrdemRetrabalho) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
        TBProducaoFases!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
        TBProducaoFases!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
        
        'Verifica se a maquina agrega custos/eficiencia na ordem
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & TBFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            If TBMaquinas!custos = True Then TBProducaoFases!custos = True Else TBProducaoFases!custos = False
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then TBProducaoFases!Valor_hs_prep = TBMaquinas!PrecoHora_Setup Else TBProducaoFases!Valor_hs_prep = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
            TBProducaoFases!Valor_hs_exec = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
        End If
        TBMaquinas.Close
        
        TBProducaoFases!IDPROCESSO = TBFases!IDPROCESSO
        TBProducaoFases!Ordem = OrdemRetrabalho
        TBProducaoFases!quantidade = QtdeOrdemRetrabalho
        TBProducaoFases!pecahora = TBFases!pecahora
        If TBFases!pecahora = True Then
            TBProducaoFases!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        Else
            If IsNull(TBFases!Execucao) = False And TBFases!Execucao <> "00:00:00" Then
                ElapsedTime (TBFases!Execucao)
                TBProducaoFases!Pcshora = 3600 / s
            End If
        End If
        TBProducaoFases!pc_te = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
        TBProducaoFases!Preparacao = IIf(IsNull(TBFases!Preparacao), "00:00:00", TBFases!Preparacao)
        TBProducaoFases!Execucao = IIf(IsNull(TBFases!Execucao), "00:00:00", TBFases!Execucao)
        TBProducaoFases!TempoPreparacao = TBFases!TempoPreparacao
        TBProducaoFases!TempoExecucao = TBFases!TempoExecucao
        TBProducaoFases!descfase = TBFases!Descricao
        If IsNull(TBFases!TESegundos) = True Or TBFases!TESegundos = "" Then
            ElapsedTime (TBProducaoFases!Execucao)
            TBProducaoFases!TESegundos = s
        Else
            TBProducaoFases!TESegundos = TBFases!TESegundos
        End If
        
        TBProducaoFases.Update
        TBFases.MoveNext
    Loop
End If

'Prazo final da OS
ProcDefinirPrazosOS OrdemRetrabalho, PrazoEntregaOrdem, False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcGravarNCOSMaq()
On Error GoTo tratar_erro

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Ordemservico_maq_utilizadas where OS = " & cmbOS & " and Maquina = '" & Cmb_maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    QTNC = 0
    Set TBProducaoFases = CreateObject("adodb.recordset")
    TBProducaoFases.Open "Select Sum(TTNC) as QTNC from CQ_NC_FABRICA where Codigo <> " & IIf(txtidos = "", 0, txtidos) & " and OS = " & cmbOS & " and Maquina = '" & Cmb_maquina & "' and (PARECERCQ = 'Rejeitar' or PARECERCQ = 'Retrabalhar' or PARECERCQ = 'Selecionar' or PARECERCQ = 'Outros' or PARECERCQ = 'Nada consta')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProducaoFases.EOF = False Then
        QTNC = IIf(IsNull(TBProducaoFases!QTNC), 0, TBProducaoFases!QTNC)
    End If
    
    If Novo_CQNC = True Then
        If Opt_rejeitar.Value = True Then TBproducao!QTNC = QTNC + IIf(txtnc = "", 0, txtnc)
        If Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True Then
            'TBproducao!QTNC = IIf(QTNC - IIf(txtnc = "", 0, txtnc) < 0, 0, QTNC - IIf(txtnc = "", 0, txtnc))
            TBproducao!QTNC = QTNC
            TBproducao!QTOK = TBproducao!QTOK + IIf(txtnc = "", 0, txtnc)
        End If
    Else
        'Atualiza dados no apontamento
        If ListaFases.SelectedItem.ListSubItems(1) <> "" And ListaFases.SelectedItem.ListSubItems(1) <> "0" Then
            QTNCAP = 0
            Set TBProducaoFases = CreateObject("adodb.recordset")
            TBProducaoFases.Open "Select Sum(TTNC) as QTNCAP from CQ_NC_FABRICA where Codigo <> " & IIf(txtidos = "", 0, txtidos) & " and idproducao = " & ListaFases.SelectedItem.ListSubItems(1) & " and (PARECERCQ = 'Rejeitar' or PARECERCQ = 'Retrabalhar' or PARECERCQ = 'Selecionar' or PARECERCQ = 'Outros' or PARECERCQ = 'Nada consta')", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducaoFases.EOF = False Then
                QTNCAP = IIf(IsNull(TBProducaoFases!QTNCAP), 0, TBProducaoFases!QTNCAP)
            End If
        End If
        
        If TBGravar!ParecerCQ = "NULL" Or TBGravar!ParecerCQ = "" Or TBGravar!ParecerCQ = "Nada consta" Then
            If Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True Or Opt_reaproveitar.Value = True Then
                TBproducao!QTOK = TBproducao!QTOK + IIf(txtnc = "", 0, txtnc)
                'TBproducao!QTNC = IIf(QTNC - IIf(txtnc = "", 0, txtnc) < 0, 0, QTNC - IIf(txtnc = "", 0, txtnc))
                TBproducao!QTNC = QTNC

                'Atualiza dados no apontamento
                If ListaFases.SelectedItem.ListSubItems(1) <> "" And ListaFases.SelectedItem.ListSubItems(1) <> "0" Then
                    Set TBProducaoFases = CreateObject("adodb.recordset")
                    TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProducaoFases.EOF = False Then
                        'TBProducaoFases!Reprovada = IIf(QTNCAP - IIf(txtnc = "", 0, txtnc) < 0, 0, QTNCAP - IIf(txtnc = "", 0, txtnc))
                        TBProducaoFases!Reprovada = QTNCAP
                        TBProducaoFases!quantidade = TBProducaoFases!quantidade + IIf(txtnc = "", 0, txtnc)
                        TBProducaoFases.Update
                        
                        ProcAtualizaMovEstoque
                    End If
                    TBProducaoFases.Close
                End If
            End If
        Else
            If TBGravar!ParecerCQ <> "Rejeitar" And TBGravar!ParecerCQ <> "Retrabalhar" And TBGravar!ParecerCQ <> "Selecionar" And TBGravar!ParecerCQ <> "Outros" And (Opt_rejeitar.Value = True Or Opt_retrabalhar.Value = True Or Opt_selecionar.Value = True Or Opt_outros.Value = True Or Opt_nada_consta.Value = True) Then
                TBproducao!QTNC = QTNC + IIf(txtnc = "", 0, txtnc)
                TBproducao!QTOK = IIf(TBproducao!QTOK - IIf(txtnc = "", 0, txtnc) < 0, 0, TBproducao!QTOK - IIf(txtnc = "", 0, txtnc))

                'Atualiza dados no apontamento
                If ListaFases.SelectedItem.ListSubItems(1) <> "" And ListaFases.SelectedItem.ListSubItems(1) <> "0" Then
                    Set TBProducaoFases = CreateObject("adodb.recordset")
                    TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProducaoFases.EOF = False Then
                        TBProducaoFases!Reprovada = QTNCAP + IIf(txtnc = "", 0, txtnc)
                        TBProducaoFases!quantidade = IIf(TBProducaoFases!quantidade - IIf(txtnc = "", 0, txtnc) < 0, 0, TBProducaoFases!quantidade - IIf(txtnc = "", 0, txtnc))
                        TBProducaoFases.Update
                        
                        ProcAtualizaMovEstoque
                    End If
                    TBProducaoFases.Close
                End If
            End If
            If TBGravar!ParecerCQ <> "Aprovado" And TBGravar!ParecerCQ <> "Aprovado c/ desvio" And TBGravar!ParecerCQ <> "Reaproveitar" And (Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True) Then
                'TBproducao!QTNC = IIf(QTNC - IIf(txtnc = "", 0, txtnc) < 0, 0, QTNC - IIf(txtnc = "", 0, txtnc))
                TBproducao!QTNC = QTNC
                TBproducao!QTOK = TBproducao!QTOK + IIf(txtnc = "", 0, txtnc)

                'Atualiza dados no apontamento
                If ListaFases.SelectedItem.ListSubItems(1) <> "" And ListaFases.SelectedItem.ListSubItems(1) <> "0" Then
                    Set TBProducaoFases = CreateObject("adodb.recordset")
                    'TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                    TBProducaoFases.Open "Select * from producaofases where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProducaoFases.EOF = False Then
                        TBProducaoFases!Reprovada = QTNCAP
                        'TBProducaoFases!Reprovada = IIf(QTNCAP - IIf(txtnc = "", 0, txtnc) < 0, 0, QTNCAP - IIf(txtnc = "", 0, txtnc))
                        TBProducaoFases!quantidade = TBProducaoFases!quantidade + IIf(txtnc = "", 0, txtnc)
                        TBProducaoFases.Update
                        
                        ProcAtualizaMovEstoque
                    End If
                    TBProducaoFases.Close
                End If
            End If
        End If
    End If

    TBproducao!Totalprod = TBproducao!QTOK + TBproducao!QTNC
    TBproducao.Update
End If
TBproducao.Close

ProcGravarNCOS cmbOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcAtualizaMovEstoque()
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from Producao where Ordem = " & TBProducaoFases!Ordem & " and Entrar_estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    LATexto = "ESTOQUE PADRÃO"
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select ELAC.Descricao FROM Estoque_Localarmazenamento_criar ELAC INNER JOIN Estoque_Localarmazenamento ELA ON ELA.idemb_locarm = ELAC.id where ELA.codinterno = '" & TBOrdem!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then LATexto = TBCFOP!Descricao
    TBCFOP.Close
    
    Permitido = False
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from estoque_controle where ID_empresa = " & TBOrdem!ID_empresa & " and desenho = '" & TBOrdem!Desenho & "' and lote = '" & TBOrdem!Ordem & "' and certificado = '" & 0 & "' and corrida = '" & 0 & "' and local_armaz = '" & LATexto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!IDEstoque & " and Lote = '" & TBOrdem!Ordem & "' and Data = '" & Format(TBProducaoFases!Data, "Short Date") & "' and Responsavel = '" & TBProducaoFases!Usuario & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            TBProduto!Entrada = TBProducaoFases!quantidade
            TBProduto!Entrada_PC = TBProducaoFases!quantidade
        
            ProcEmpenharREAutomOrdem TBEstoque!IDEstoque, TBProduto!Entrada, TBEstoque!LOTE, TBProduto!Data, TBProduto!Responsavel, TBOrdem!Desenho, False
        
            Qtde = TBOrdem!Quant
            Entrada = TBProduto!Entrada
            If Entrada >= Qtde Then
                TBProduto!Operacao = "ENTRADA_ORDEM"
                TBEstoque!status = "ENTRADA_ORDEM"
            ElseIf Entrada < Qtde Then
                    TBProduto!Operacao = "ENTRADA_ORDEM_PARCIAL"
                    TBEstoque!status = "ENTRADA_ORDEM_PARCIAL"
            End If
        
            TBProduto.Update
            
            Set TBCorretiva = CreateObject("adodb.recordset")
            TBCorretiva.Open "Select Sum(Entrada) as Qtde, Sum(Saida) as qtdeliberada from estoque_movimentacao where idestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBCorretiva.EOF = False Then
                Qtde = IIf(IsNull(TBCorretiva!Qtde), 0, TBCorretiva!Qtde) - IIf(IsNull(TBCorretiva!qtdeliberada), 0, TBCorretiva!qtdeliberada)
            End If
            TBCorretiva.Close
            
            TBEstoque!estoque_real = Qtde
            TBEstoque!estoque_real_PC = Format(Qtde, "###,##0.0000")
            TBEstoque!estoque_venda = Qtde
            TBEstoque.Update
        
            'Exclui o empenho no produto em estoque para o pedido
            Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBProduto!IDEstoque
            ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem TBProduto!LOTE, TBProduto!Desenho, TBEstoque!estoque_real, TBProduto!IDEstoque
        
            qtdeliberada = 0
            QtdeSaida = 0
            QtdeEstoque = 0
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from projproduto where desenho = '" & TBOrdem!Desenho & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = True Then
                    TBEstoque.Delete
                Else
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Sum(ISNULL(Entrada, 0)) as qtdeliberada, Sum(ISNULL(Saida, 0)) as QtdeSaida from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        qtdeliberada = IIf(IsNull(TBProduto!qtdeliberada), 0, TBProduto!qtdeliberada)
                        QtdeSaida = IIf(IsNull(TBProduto!QtdeSaida), 0, TBProduto!QtdeSaida)
                        NovoValor = Replace(qtdeliberada - QtdeSaida, ",", ".")
                        Conexao.Execute "Update Estoque_movimentacao Set estoque_venda = " & NovoValor & " where IDestoque = " & TBEstoque!IDEstoque & " and Lote = '" & TBOrdem!Ordem & "' and Data = '" & Format(TBProducaoFases!Data, "Short Date") & "' and Responsavel = '" & TBProducaoFases!Usuario & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')"
                    End If
                End If
            End If
            TBEstoque.Update
        End If
        TBProduto.Close
    End If
    TBEstoque.Close
    
    'Corrige valor do estoque
                                              'ORDEM         QTDE. PREVISTA                                QTDE. OK                                              QT. PROD.(OK+NC)                                                                                         CUSTO LOTE                                        CUSTO PEÇA                                CUSTO TERCEIROS                                       CUSTO MATERIAL                                      CUSTO OUTROS                                        ORDEM CONSIGNADA
    Valor_Produto = FunCalculaValorUnitOrdem(TBOrdem!Ordem, IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd) + IIf(IsNull(TBOrdem!QuantNC), 0, TBOrdem!QuantNC), IIf(IsNull(TBOrdem!CTTReal), 0, TBOrdem!CTTReal), IIf(IsNull(TBOrdem!CPR), 0, TBOrdem!CPR), IIf(IsNull(TBOrdem!CTServico), 0, TBOrdem!CTServico), IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial), IIf(IsNull(TBOrdem!CTOutras), 0, TBOrdem!CTOutras), TBOrdem!Consignacao)
    NovoValor = Replace(Valor_Produto, ",", ".")
    Conexao.Execute "Update Estoque_Controle set valor_unitario = " & NovoValor & " where Lote = '" & TBOrdem!Ordem & "' and Desenho = '" & TBOrdem!Desenho & "'"
    Conexao.Execute "Update Estoque_Controle set Valor_Total = ROUND(valor_unitario * estoque_real, 2) where Lote = '" & TBOrdem!Ordem & "' and Desenho = '" & TBOrdem!Desenho & "'"
    Conexao.Execute "Update EM Set VlrUnit = " & NovoValor & " from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EM.IdEstoque = EC.IdEstoque where EC.Lote = '" & TBOrdem!Ordem & "' and EM.Desenho = '" & TBOrdem!Desenho & "' and (EM.Operacao = 'ENTRADA_ORDEM' Or EM.Operacao = 'ENTRADA_ORDEM_PARCIAL')"
    Conexao.Execute "Update EM Set VlrTotal = ROUND(VlrUnit * Entrada, 2) from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EM.IdEstoque = EC.IdEstoque where EC.Lote = '" & TBOrdem!Ordem & "' and EM.Desenho = '" & TBOrdem!Desenho & "' and (EM.Operacao = 'ENTRADA_ORDEM' Or EM.Operacao = 'ENTRADA_ORDEM_PARCIAL')"
    Conexao.Execute "Update CC Set CC.Valor = EM.VlrTotal from (CC_realizado CC INNER JOIN Estoque_movimentacao EM ON CC.ID_estoque = EM.Idoperacao) INNER JOIN Estoque_controle EC ON EC.IdEstoque = EM.IdEstoque where EC.Lote = '" & TBOrdem!Ordem & "' and EC.Desenho = '" & TBOrdem!Desenho & "' and (EM.Operacao = 'ENTRADA_ORDEM' Or EM.Operacao = 'ENTRADA_ORDEM_PARCIAL')"
    Conexao.Execute "Update EM set EM.VlrUnit = EM1.VlrUnit from Estoque_movimentacao EM INNER JOIN Estoque_movimentacao EM1 on EM1.IdEstoque = EM.IdEstoque where EM.Lote = '" & TBOrdem!Ordem & "' and EM.Desenho = '" & TBOrdem!Desenho & "' and EM.Saida > 0 and EM1.Entrada > 0"
    Conexao.Execute "Update EC set EC.valor_unitario = EM1.VlrUnit from Estoque_Controle EC INNER JOIN Estoque_movimentacao EM ON EM.IdEstoque = EC.IdEstoque INNER JOIN Estoque_movimentacao EM1 on EM1.IdEstoque = EM.IdEstoque where EM.Lote = '" & TBOrdem!Ordem & "' and EM.Desenho = '" & TBOrdem!Desenho & "' and EM.Saida > 0 and EM1.Entrada > 0"
    Conexao.Execute "Update Estoque_movimentacao set VlrTotal = ROUND(VlrUnit * Saida, 2) where Lote = '" & TBOrdem!Ordem & "' and Desenho = '" & TBOrdem!Desenho & "' and Saida > 0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcGravarNCOS(OS As Long)
On Error GoTo tratar_erro

'Atualiza qtde. NC e produzida na OS
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(QTOK) as QTPC, Sum(QTNC) as QTNC, Sum(Totalprod) as QTLOTE from Ordemservico_maq_utilizadas where OS = " & OS, Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    NovoValor = Replace(IIf(IsNull(TBproducao!QTPC), 0, TBproducao!QTPC), ",", ".")
    NovoValor1 = Replace(IIf(IsNull(TBproducao!QTNC), 0, TBproducao!QTNC), ",", ".")
    NovoValor2 = Replace(IIf(IsNull(TBproducao!QTLOTE), 0, TBproducao!QTLOTE), ",", ".")
    Conexao.Execute "Update Ordemservico Set QTOK = " & NovoValor & ", QTNC = " & NovoValor1 & ", Totalprod = " & NovoValor2 & " where IDProducao = " & OS
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcGravarNCOrdem(Ordem As Long)
On Error GoTo tratar_erro

TotalNC = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select SUM(ISNULL(TTNC, 0)) as TotalNC from CQ_NC_FABRICA where Ordem = " & Ordem & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
              
    TotalNC = IIf(IsNull(TBAbrir!TotalNC), 0, TBAbrir!TotalNC)
    
    
    
End If
TBAbrir.Close
Conexao.Execute "Update Producao Set quantNC = " & Replace(TotalNC, ",", ".") & " where Ordem = " & Ordem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExcluirNCOSMaq(OS As Long)
On Error GoTo tratar_erro

If IsNull(TBFI!IDProducao) = True Or TBFI!IDProducao = 0 Then
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select * from Ordemservico_maq_utilizadas where OS = " & OS, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        If TBFI!ParecerCQ = "Rejeitar" Then TBproducao!QTNC = IIf(TBproducao!QTNC - IIf(IsNull(TBFI!TTNC), 0, TBFI!TTNC) < 0, 0, TBproducao!QTNC - IIf(IsNull(TBFI!TTNC), 0, TBFI!TTNC))
        If TBFI!ParecerCQ = "Aprovado" Or TBFI!ParecerCQ = "Aprovado c/ desvio" Then TBproducao!QTOK = IIf(TBproducao!QTOK - IIf(IsNull(TBFI!TTNC), 0, TBFI!TTNC) < 0, 0, TBproducao!QTOK - IIf(IsNull(TBFI!TTNC), 0, TBFI!TTNC))
        TBproducao!Totalprod = TBproducao!QTOK + TBproducao!QTNC
        TBproducao.Update
    End If
    TBproducao.Close
    ProcGravarNCOS TBFI!OS
    ProcGravarNCOrdem TBFI!Ordem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtaprovadas_LostFocus()
On Error GoTo tratar_erro

If txtaprovadas.Text <> "" Then
    VerifNumero = txtaprovadas.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtaprovadas.Text = ""
        txtaprovadas.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtdata_Change()
On Error GoTo tratar_erro

ProcVerificaTurno

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtHora_Change()
On Error GoTo tratar_erro

ProcVerificaTurno

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerificaTurno()
On Error GoTo tratar_erro

Txt_turno = 0
TempoInicio = 0
TempoFinal = 0
Dataini = txtData
ProcVerificaDia
Dataini = txthora
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from CadMaqturnos where maquina = '" & Cmb_maquina & "' and diasemana = '" & Diasemana & "' order by diasemana,turno", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Do While TBMaquinas.EOF = False
        If IsNull(TBMaquinas!Inicioturno) = False Then
            TempoInicio = TBMaquinas!Inicioturno
            TempoFinal = TBMaquinas!finalturno
            If TempoInicio > TempoFinal Then
                Dataini = txtData & " " & Dataini
                TempoInicio = txtData & " " & TempoInicio
                TempoFinal = txtData & "  " & TempoFinal
                TempoInicio = TempoInicio - 1
                TempoFinal = TempoFinal + 1
            End If
            Select Case TBMaquinas!Turno
                Case 1:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Txt_turno = 1
                        GoTo Sair
                    End If
                Case 2:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Txt_turno = 2
                        GoTo Sair
                    End If
                Case 3:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Txt_turno = 3
                        GoTo Sair
                    End If
                Case 4:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Txt_turno = 4
                        GoTo Sair
                    End If
            End Select
        End If
        TBMaquinas.MoveNext
    Loop
End If
Sair:
    TBMaquinas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerificaDia()
On Error GoTo tratar_erro

Diasemana = Weekday(Dataini)
Select Case Diasemana
    Case 1: Diasemana = "Domingo"
    Case 2: Diasemana = "Segunda"
    Case 3: Diasemana = "Terça"
    Case 4: Diasemana = "Quarta"
    Case 5: Diasemana = "Quinta"
    Case 6: Diasemana = "Sexta"
    Case 7: Diasemana = "Sabado"
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtnc_Change()
On Error GoTo tratar_erro

'If txtnc.Text <> "" Then
'    VerifNumero = txtnc.Text
'    ProcVerificaNumero
'    If VerifNumero = False Then
'        txtnc.Text = ""
'        txtnc.SetFocus
'        Exit Sub
'    End If
'End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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

Private Sub txtOF_Change()
On Error GoTo tratar_erro

If txtof.Text <> "" Then
    VerifNumero = txtof.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtof.Text = ""
        txtof.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtof_LostFocus()
On Error GoTo tratar_erro

ProcLimpaCamposOrdem
ProcCarregaComboOS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLimpaCamposOrdem()
On Error GoTo tratar_erro

Cmb_maquina.ListIndex = -1
txtFase.Text = ""
txtdesenho.Text = ""
txtreferencia = ""
txtdescricao.Text = ""
txtLote.Text = ""
txtaprovadas.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaComboOS()
On Error GoTo tratar_erro

With cmbOS
    .Clear
    If txtof <> "" Then
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select IDProducao from ordemservico where Ordem = " & txtof, Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            Do While TBproducao.EOF = False
                .AddItem TBproducao!IDProducao
                TBproducao.MoveNext
            Loop
        End If
        TBproducao.Close
    End If
End With

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

Private Sub USButton1_Click()

End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcReposicaoRetrabalho
    Case 7: procOsRetrabalho
    Case 8: procOrdemRetrabalho
    Case 9: ProcRNC
    Case 10: procDisposicao
    Case 11: ProcAtualizar
    'Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procOsRetrabalho()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Acao = "criar OS de retrabalho"
If txtidos = "" Then
    NomeCampo = "a não conformidade"
    ProcVerificaAcao
    Exit Sub
End If

'Retrabalho
If Opt_retrabalhar = False Then
    USMsgBox "Não é permitido criar OS de retrabalho, pois a disposição é diferente de retrabalhar.", vbInformation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBOSC = CreateObject("adodb.recordset")
TBOSC.Open "Select Ordem from Ordemservico where Ordem = " & txtof & " and fase = " & txtFase & " and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockReadOnly
If TBOSC.EOF = False Then
    USMsgBox "Não é permitido criar OS de retrabalho, pois já existe OS de retrabalho criada para essa não conformidade.", vbInformation, "CAPRIND v5.0"
    TBOSC.Close
    Exit Sub
End If
TBOSC.Close

If USMsgBox("Deseja emitir uma nova OS de retrabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
    QuantSolicitado = 0
    Retrabalho = InputBox("Favor informar a quantidade de retrabalho?")
    If IsNumeric(Retrabalho) = True Then
        QuantSolicitado = Retrabalho
    Else
        USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    If QuantSolicitado <> 0 Then ProcCriarOSRetrabalho
    USMsgBox ("OS de retrabalho criada com sucesso."), vbInformation, "CAPRIND v5.0"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procDisposicao()
On Error GoTo tratar_erro

Permitido = False
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) não conformidade(s) antes de alterar a disposição."), vbExclamation, "CAPRIND v5.0"
Else
    frmcqnc_disposicao.Show 1
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcRNC()
On Error GoTo tratar_erro

Qtde = 0
Qtd = 0
Permitido = False
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            quantidade = .ListItems(InitFor)
            Qtde = Qtde + .ListItems(InitFor).ListSubItems(7)
            Permitido = True
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) não conformidade(s) antes de criar RNC."), vbExclamation, "CAPRIND v5.0"
Else
    RNC_Controle_Medicao = False
    RNC_Inspecao_Recebimento = False
    RNC_Nao_Conformidade = True
    RNC_Solicitacao_Desvio = False
    Sit_REG = 2
    frmQualidade_RNC.Show
    
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 Then
        If ListaFases.ListItems.Count <> 0 Then
            ListaFases.SelectedItem = ListaFases.ListItems(CodigoLista)
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcImprimirRel_Variavel(FormulaRel As String)
On Error GoTo tratar_erro

If ListaFases.ListItems.Count = 0 Then Exit Sub

ProcExcluirDadosProducaoRelatorios
Set TBRecebidos = CreateObject("adodb.recordset")
TBRecebidos.Open StrSql_CQ_NC_FIltro, Conexao, adOpenKeyset, adLockReadOnly
Do While TBRecebidos.EOF = False
    Set TBGravar = CreateObject("adodb.recordset")
    
    'Conexao.Execute "DELETE from Producao_Relatorios where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'"
    
    TBGravar.Open "SELECT * FROM Producao_Relatorios WHERE Execucaoprev = '" & TBRecebidos!Desenho & "' and Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
    DataFiltro = ""
    If dataTipo_CQNC = 1 Or dataTipo_CQNC = 2 Or dataTipo_CQNC = 3 Then
        If dataTipo_CQNC = 1 Then
            CampoData = "Data"
        ElseIf dataTipo_CQNC = 2 Then
            CampoData = "DataEmissao"
        Else
            CampoData = "DataConclusao"
        End If
        DataFiltro = " and " & CampoData & " Between '" & Format(dataTipo_CQNC_DE, "Short Date") & "' And '" & Format(dataTipo_CQNC_Ate, "Short Date") & "'"
    End If
    
    If pesquisaPorOrdem = True Then
        selectfiltro = "quant"
        wherefiltro = "ordem = " & TBRecebidos!Ordem & " AND desenho = '" & TBRecebidos!Desenho
    Else
        selectfiltro = "SUM(quant) as quant"
        wherefiltro = "desenho = '" & TBRecebidos!Desenho
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "SELECT " & selectfiltro & " FROM ESPLENDOR_CQNC_Totalinspecionado where " & wherefiltro & "'" & DataFiltro, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        TBGravar!Execucaoprev = TBRecebidos!Desenho
        TBGravar!OS = TBGravar!OS + TBAbrir!Quant
        TBGravar!Responsavel = pubUsuario
        TBGravar!Modulo = Formulario
        TBGravar.Update
    End If
    TBAbrir.Close

    TBRecebidos.MoveNext
Loop
TBRecebidos.Close

Set Report = crAPP.OpenReport(Localrel & "\Personalizados\" & NomeRel)
'Login SQL
Contador = Report.Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
Loop
ProcVerifSubReport FormulaRel_CQ_NC
'procVerifSubReport_Esplendor

frmimprimir.CrystalActiveXReportViewer1.ReportSource = Report
Report.FormulaSyntax = crCrystalSyntaxFormula
Report.RecordSelectionFormula = FormulaRel

Report.ParameterFields(1).AddCurrentValue (dataTipo_CQNC)
Report.ParameterFields(2).AddCurrentValue (pubUsuario)
Report.ParameterFields(3).AddCurrentValue (Formulario)

frmimprimir.CrystalActiveXReportViewer1.ViewReport
frmimprimir.Show 1
2:
    Set Report = Nothing
    Set crAPP = Nothing

Exit Sub
tratar_erro:
    If Err.Number = "-2147206461" Then
        USMsgBox ("Não foi encontrado o relatório " & NomeRel & " na pasta " & LocalrelNovo), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    If Err.Number = "-2147483638" Then
        USMsgBox ("Não foi possível visualizar o relatório, favor reiniciar o sistema."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procOrdemRetrabalho()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

'Soma quantidades
QtdeOrdemRetrabalho = 0
Permitido1 = True
clienteordem = ""
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            QtdeOrdemRetrabalho = QtdeOrdemRetrabalho + 1
            
            'Verifica se tem cliente diferente, se tiver ele não salva o clinte na ordem de retrabalho
            If Permitido1 = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select idcliente from producao where ordem = " & .ListItems(InitFor).ListSubItems(2), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If clienteordem = "" Then
                        clienteordem = TBContas!IDCliente
                    ElseIf clienteordem <> TBContas!IDCliente Then
                        Permitido1 = False
                    End If
                End If
                TBContas.Close
            End If
        End If
    Next InitFor
End With

Permitido = False
OrdemRetrabalho = 0
With ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente criar uma nova ordem de retabalho?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                
Mensagem:
                DataTexto = InputBox("Favor informar o prazo de entrega da ordem de retrabalho.")
                If DataTexto = "" Then Exit Sub
                If IsDate(DataTexto) = False Then
                    USMsgBox ("Esta data não é válida."), vbExclamation, "CAPRIND v5.0"
                    GoTo Mensagem
                End If
                PrazoEntregaOrdem = DataTexto
                
                ProcCriarOrdemRetrabalho .ListItems(InitFor).ListSubItems(2)
            End If
            Permitido = True
            
            'Vincular pedido interno
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from producao_pedidos where Ordem = " & .ListItems(InitFor).ListSubItems(2), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from producao_pedidos where Ordem = " & OrdemRetrabalho & " and IDCarteira = " & TBAbrir!IDcarteira, Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then TBGravar.AddNew
                    TBGravar!Ordem = OrdemRetrabalho
                    TBGravar!IDcarteira = TBAbrir!IDcarteira
                    TBGravar.Update
                    
                    Conexao.Execute "Update vendas_carteira Set Tem_ordem = 'True' where Codigo = " & TBAbrir!IDcarteira
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            
            Conexao.Execute "Update CQ_NC_FABRICA Set OrdemRetrabalho = " & OrdemRetrabalho & " where codigo = " & .ListItems(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) não conformidade(s) antes de criar ordem de retrabalho."), vbExclamation, "CAPRIND v5.0"
Else
    procAtualizaPedido
    USMsgBox ("Ordem de retrabalho criada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame1.Enabled = False
    Frame3.Enabled = False
    Novo_CQNC = False
    ProcCarregaLista (1)
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function funcCriarNumeroOrdem() As Long
On Error GoTo tratar_erro

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select Ordem from producao order by Ordem desc", Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = True Then ProcCriarNumeroOrdem = 1 Else funcCriarNumeroOrdem = TBContas!Ordem + 1
TBContas.Close

VerifOrdem:
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Ordem from producao where Ordem = " & funcCriarNumeroOrdem, Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        funcCriarNumeroOrdem = funcCriarNumeroOrdem + 1
        GoTo VerifOrdem
    End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub procAtualizaPedido()
On Error GoTo tratar_erro

'Vincular pedido interno
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from producao_pedidos where Ordem = " & OrdemRetrabalho, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    Qtde = TBAbrir.RecordCount
    Qtd = QtdeOrdemRetrabalho / Qtde
    Conexao.Execute "Update producao_pedidos Set Qtde_empenho = " & Qtd & " where ordem = " & OrdemRetrabalho
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procVerifSubReport_Esplendor()
On Error GoTo tratar_erro

SubReportRel = "ordens"
Contador = Report.OpenSubreport(SubReportRel).Database.Tables.Count
Do While Contador > 0
    Set DBTable = Report.OpenSubreport(SubReportRel).Database.Tables(Contador)
    ProcLogonBDSQL
    Contador = Contador - 1
    
    'Coloca a formula no subreport
    Report.OpenSubreport(SubReportRel).FormulaSyntax = crCrystalSyntaxFormula
    Report.OpenSubreport(SubReportRel).RecordSelectionFormula = FormulaRel_CQ_NC + " and {CQ_NC_FABRICA.obsFab} = {?Pm-CQ_NC_FABRICA.obsFab} and {Producao.desenho} = {?Pm-Producao.desenho}"
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
