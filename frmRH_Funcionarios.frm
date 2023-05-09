VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRH_Funcionarios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "RH - Cadastro de funcionários"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   3495
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
   ForeColor       =   &H00000000&
   Icon            =   "frmRH_Funcionarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
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
      ItemData        =   "frmRH_Funcionarios.frx":1042
      Left            =   240
      List            =   "frmRH_Funcionarios.frx":1044
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1695
      Width           =   5835
   End
   Begin MSComctlLib.ListView Lista_doc 
      Height          =   6105
      Left            =   60
      TabIndex        =   74
      Top             =   3600
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10769
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Caminho"
         Object.Width           =   25576
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
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   119
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
         ItemData        =   "frmRH_Funcionarios.frx":1046
         Left            =   6960
         List            =   "frmRH_Funcionarios.frx":1050
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   4
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   8
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmRH_Funcionarios.frx":1068
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
         TabIndex        =   7
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmRH_Funcionarios.frx":480C
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
         TabIndex        =   5
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
         TabIndex        =   6
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmRH_Funcionarios.frx":8315
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
         TabIndex        =   9
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmRH_Funcionarios.frx":C404
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
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   142
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação da lista"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   5610
         TabIndex        =   123
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
         Top             =   240
         Width           =   1095
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   114
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
   Begin DrawSuite2022.USToolBar USToolBar3 
      Height          =   975
      Left            =   55
      TabIndex        =   132
      Top             =   330
      Visible         =   0   'False
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   44
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
      ButtonLeft3     =   83
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
   End
   Begin MSComctlLib.ListView ListaFerias 
      Height          =   7610
      Left            =   60
      TabIndex        =   50
      Top             =   2140
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   13414
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
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. início férias"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. fim férias"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Dias"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Dt. início aquisitivo"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Dt. fim aquisitivo"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView ListaCursos 
      Height          =   7035
      Left            =   60
      TabIndex        =   41
      Top             =   2745
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12409
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
         Text            =   "Curso"
         Object.Width           =   12965
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Duração"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Instituição educacional"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Nível"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView ListaAtestados 
      Height          =   7605
      Left            =   60
      TabIndex        =   64
      Top             =   2145
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   13414
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsavel"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Dt. validade"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Caminho do atestado"
         Object.Width           =   13062
      EndProperty
   End
   Begin MSComctlLib.ListView ListaObs 
      Height          =   7610
      Left            =   55
      TabIndex        =   67
      Top             =   2140
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   13414
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Observações"
         Object.Width           =   23548
      EndProperty
   End
   Begin MSComctlLib.ListView ListaAumentos 
      Height          =   7035
      Left            =   60
      TabIndex        =   45
      Top             =   2745
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12409
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Observações"
         Object.Width           =   20902
      EndProperty
   End
   Begin MSComctlLib.ListView ListaSindicato 
      Height          =   7035
      Left            =   60
      TabIndex        =   54
      Top             =   2745
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12409
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Observações"
         Object.Width           =   20902
      EndProperty
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5085
      Left            =   60
      TabIndex        =   1
      Top             =   4035
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8969
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
         Object.Tag             =   "N"
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nome"
         Object.Width           =   13759
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Telefone"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Validado"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Frame Frame2 
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
      Height          =   2715
      Left            =   60
      TabIndex        =   147
      Top             =   1320
      Width           =   15200
      Begin VB.TextBox txtPIS 
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
         Left            =   4140
         MaxLength       =   11
         TabIndex        =   166
         ToolTipText     =   "PIS."
         Top             =   1650
         Width           =   1245
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
         MaxLength       =   10
         TabIndex        =   165
         ToolTipText     =   "Código."
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtNome 
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
         Left            =   870
         MaxLength       =   100
         TabIndex        =   164
         ToolTipText     =   "Nome."
         Top             =   1020
         Width           =   3885
      End
      Begin VB.TextBox txtRG 
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
         Left            =   8970
         MaxLength       =   30
         TabIndex        =   163
         ToolTipText     =   "Numero do RG."
         Top             =   1020
         Width           =   1185
      End
      Begin VB.TextBox txtReservista 
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
         TabIndex        =   162
         ToolTipText     =   "Reservista."
         Top             =   1650
         Width           =   1245
      End
      Begin VB.TextBox txtEleitor 
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
         Left            =   11535
         MaxLength       =   50
         TabIndex        =   161
         ToolTipText     =   "Título eleitor."
         Top             =   1020
         Width           =   1995
      End
      Begin VB.TextBox txtCTPS_N 
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   160
         ToolTipText     =   "N° do CTPS."
         Top             =   1650
         Width           =   1305
      End
      Begin VB.TextBox txtCTPS_serie 
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
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   159
         ToolTipText     =   "Série do CTPS."
         Top             =   1650
         Width           =   1365
      End
      Begin VB.TextBox txtPai 
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
         Left            =   9465
         MaxLength       =   100
         TabIndex        =   158
         ToolTipText     =   "Nome do pai."
         Top             =   1650
         Width           =   4065
      End
      Begin VB.TextBox txtMae 
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
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   157
         ToolTipText     =   "Nome da mãe."
         Top             =   1650
         Width           =   4055
      End
      Begin VB.ComboBox cmbSexo 
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
         ItemData        =   "frmRH_Funcionarios.frx":FC90
         Left            =   6300
         List            =   "frmRH_Funcionarios.frx":FC9D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   156
         ToolTipText     =   "Sexo."
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txtTelefone 
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
         MaxLength       =   30
         TabIndex        =   155
         ToolTipText     =   "N° do(s) telefone(s)."
         Top             =   2280
         Width           =   2705
      End
      Begin VB.TextBox txtEndereco 
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
         MaxLength       =   100
         TabIndex        =   154
         ToolTipText     =   "Endereço."
         Top             =   2280
         Width           =   6015
      End
      Begin VB.ComboBox txtCivil 
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
         ItemData        =   "frmRH_Funcionarios.frx":FCB8
         Left            =   7440
         List            =   "frmRH_Funcionarios.frx":FCD1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   153
         ToolTipText     =   "Estado cívil."
         Top             =   1020
         Width           =   1515
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
         ItemData        =   "frmRH_Funcionarios.frx":FD25
         Left            =   8925
         List            =   "frmRH_Funcionarios.frx":FD27
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   152
         ToolTipText     =   "Centro de custo."
         Top             =   2280
         Width           =   4605
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
         Left            =   10195
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         ToolTipText     =   "Data e hora da validação."
         Top             =   390
         Width           =   1635
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
         Left            =   11850
         Locked          =   -1  'True
         TabIndex        =   150
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pela validação."
         Top             =   390
         Width           =   3150
      End
      Begin VB.TextBox txtData 
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
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   149
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   7025
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   3160
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   10170
         TabIndex        =   167
         ToolTipText     =   "Número do CPF."
         Top             =   1020
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNascimento 
         Height          =   315
         Left            =   4770
         TabIndex        =   168
         ToolTipText     =   "Data de nascimento."
         Top             =   1020
         Width           =   1135
         _ExtentX        =   2011
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
      Begin DrawSuite2022.USButton Cmd_localizar_foto 
         Height          =   300
         Left            =   13650
         TabIndex        =   169
         Top             =   2310
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         Caption         =   "Localizar foto"
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
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome*"
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
         Left            =   2520
         TabIndex        =   191
         Top             =   810
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10695
         TabIndex        =   190
         Top             =   810
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   9450
         TabIndex        =   189
         Top             =   810
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIS"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4642
         TabIndex        =   188
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6675
         TabIndex        =   187
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado cívil"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7800
         TabIndex        =   186
         Top             =   810
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone(s)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7215
         TabIndex        =   185
         Top             =   2055
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2850
         TabIndex        =   184
         Top             =   2055
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Título eleitor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12090
         TabIndex        =   183
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reservista"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   182
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. nascimento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4792
         TabIndex        =   181
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do pai"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   11055
         TabIndex        =   180
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da mãe"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6940
         TabIndex        =   179
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód.*"
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
         TabIndex        =   178
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CTPS  N°                 Série"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1845
         TabIndex        =   177
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Image imgCalendario 
         Height          =   360
         Left            =   5910
         Picture         =   "frmRH_Funcionarios.frx":FD29
         Stretch         =   -1  'True
         ToolTipText     =   "Abrir calendário."
         Top             =   990
         Width           =   330
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
         Left            =   2850
         TabIndex        =   176
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10725
         TabIndex        =   175
         Top             =   2055
         Width           =   1155
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Responsável pela validação"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12435
         TabIndex        =   174
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data/hora validação"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10285
         TabIndex        =   173
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   6345
         TabIndex        =   172
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   8143
         TabIndex        =   171
         Top             =   180
         Width           =   915
      End
      Begin VB.Image Foto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   13650
         Stretch         =   -1  'True
         ToolTipText     =   "Foto."
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackColor       =   &H00EAEAEA&
         BackStyle       =   0  'Transparent
         Caption         =   " (.JPG | .BMP) TAMANHO: ALT. 89 PX x LARG. 89 PX"
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
         Height          =   645
         Left            =   13770
         TabIndex        =   170
         Top             =   1290
         Width           =   1125
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
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
      TabPicture(0)   =   "frmRH_Funcionarios.frx":101AC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Dados adicionais"
      TabPicture(1)   =   "frmRH_Funcionarios.frx":101C8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(1)=   "USImageList2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Cursos"
      TabPicture(2)   =   "frmRH_Funcionarios.frx":101E4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "USImageList3"
      Tab(2).Control(1)=   "Frame11"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Aumentos"
      TabPicture(3)   =   "frmRH_Funcionarios.frx":10200
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Férias"
      TabPicture(4)   =   "frmRH_Funcionarios.frx":1021C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame13"
      Tab(4).Control(1)=   "Frame14"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Sindicato"
      TabPicture(5)   =   "frmRH_Funcionarios.frx":10238
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame15"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Atestados"
      TabPicture(6)   =   "frmRH_Funcionarios.frx":10254
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Observações"
      TabPicture(7)   =   "frmRH_Funcionarios.frx":10270
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame16"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Documentos"
      TabPicture(8)   =   "frmRH_Funcionarios.frx":1028C
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "txtID_doc"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Frame4"
      Tab(8).ControlCount=   2
      Begin VB.TextBox txtID_doc 
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
         Left            =   -65160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   146
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2490
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2265
         Left            =   -74945
         TabIndex        =   143
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_visualizar_doc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmRH_Funcionarios.frx":102A8
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_obs_doc 
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
            Height          =   1095
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            ToolTipText     =   "Observação."
            Top             =   1020
            Width           =   14835
         End
         Begin VB.TextBox Txt_caminho_doc 
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            ToolTipText     =   "Caminho."
            Top             =   390
            Width           =   11385
         End
         Begin VB.CommandButton Cmd_localizar_doc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmRH_Funcionarios.frx":1086A
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Localizar arquivo (F2)"
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtData_doc 
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
            MaxLength       =   25
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   855
         End
         Begin VB.TextBox txtResponsavel_doc 
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   69
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   7155
            TabIndex        =   145
            Top             =   810
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"frmRH_Funcionarios.frx":1096C
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   450
            TabIndex        =   144
            Top             =   180
            Width           =   9120
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   1965
         Left            =   -74945
         TabIndex        =   80
         Top             =   1305
         Width           =   15200
         Begin VB.CommandButton cmdFornecedor 
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
            Left            =   14700
            Picture         =   "frmRH_Funcionarios.frx":10A20
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Localizar agência."
            Top             =   1500
            Width           =   315
         End
         Begin VB.TextBox txtIDFornecedor 
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
            MaxLength       =   40
            TabIndex        =   28
            ToolTipText     =   "Código da agência."
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtFornecedor 
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
            Left            =   930
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Agência de emprego."
            Top             =   1500
            Width           =   13755
         End
         Begin VB.TextBox txtDependente 
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
            MaxLength       =   20
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "N° de dependentes."
            Top             =   390
            Width           =   795
         End
         Begin VB.CommandButton cmdDependentes 
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
            Left            =   990
            Picture         =   "frmRH_Funcionarios.frx":10B22
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Localizar dependentes."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtFuncao 
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Função do funcionário."
            Top             =   390
            Width           =   3555
         End
         Begin VB.CommandButton cmdSetor 
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
            Left            =   12270
            Picture         =   "frmRH_Funcionarios.frx":10C24
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Localizar setores."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtSetor 
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
            Left            =   8970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Setor."
            Top             =   390
            Width           =   3285
         End
         Begin VB.TextBox txtSalario 
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
            Left            =   10110
            MaxLength       =   50
            TabIndex        =   25
            ToolTipText     =   "Salário atual."
            Top             =   960
            Width           =   1965
         End
         Begin VB.CommandButton cmdDivisao 
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
            Left            =   14700
            Picture         =   "frmRH_Funcionarios.frx":10D26
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Localizar divisões."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtDivisao 
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
            Left            =   12690
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Divisão."
            Top             =   390
            Width           =   1995
         End
         Begin VB.TextBox txtConta_corrente 
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
            Left            =   6870
            MaxLength       =   50
            TabIndex        =   24
            ToolTipText     =   "Conta corrente."
            Top             =   960
            Width           =   3225
         End
         Begin VB.ComboBox cmbSituacao 
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
            ItemData        =   "frmRH_Funcionarios.frx":10E28
            Left            =   12090
            List            =   "frmRH_Funcionarios.frx":10E3B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Situação."
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtTurno 
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
            MaxLength       =   40
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Turno."
            Top             =   960
            Width           =   4575
         End
         Begin VB.CommandButton cmdTurno 
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
            Left            =   4770
            Picture         =   "frmRH_Funcionarios.frx":10E69
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar turnos."
            Top             =   960
            Width           =   315
         End
         Begin VB.CommandButton cmdFuncao 
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
            Left            =   5010
            Picture         =   "frmRH_Funcionarios.frx":10F6B
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Localizar funções."
            Top             =   390
            Width           =   315
         End
         Begin VB.ComboBox txtTipo 
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
            ItemData        =   "frmRH_Funcionarios.frx":1106D
            Left            =   5460
            List            =   "frmRH_Funcionarios.frx":1107A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Tipo."
            Top             =   390
            Width           =   1905
         End
         Begin MSMask.MaskEdBox txtHorario_inicio 
            Height          =   315
            Left            =   5190
            TabIndex        =   22
            ToolTipText     =   "Horário de entrada."
            Top             =   960
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtHorario_fim 
            Height          =   315
            Left            =   6135
            TabIndex        =   23
            ToolTipText     =   "Horário de saida."
            Top             =   960
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtData_adm 
            Height          =   315
            Left            =   7380
            TabIndex        =   15
            ToolTipText     =   "Data de nascimento"
            Top             =   390
            Width           =   1125
            _ExtentX        =   1984
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
         Begin MSMask.MaskEdBox txtdata_desligado 
            Height          =   315
            Left            =   13560
            TabIndex        =   27
            ToolTipText     =   "Data do afastamento."
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
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
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agência de emprego"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7072
            TabIndex        =   116
            Top             =   1290
            Width           =   1470
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° depen."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   210
            TabIndex        =   112
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Função"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2955
            TabIndex        =   91
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6262
            TabIndex        =   90
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. admissão"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7477
            TabIndex        =   89
            Top             =   180
            Width           =   930
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10417
            TabIndex        =   88
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salário"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10852
            TabIndex        =   87
            Top             =   750
            Width           =   480
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ag./CC"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8220
            TabIndex        =   86
            Top             =   750
            Width           =   525
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Divisão"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13432
            TabIndex        =   85
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situação"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12510
            TabIndex        =   84
            Top             =   750
            Width           =   615
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. afastado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13560
            TabIndex        =   83
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Turno"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2257
            TabIndex        =   82
            Top             =   750
            Width           =   420
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
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
            Height          =   195
            Left            =   5970
            TabIndex        =   81
            Top             =   1020
            Width           =   90
         End
         Begin VB.Image imgCalendario1 
            Height          =   360
            Left            =   8520
            Picture         =   "frmRH_Funcionarios.frx":11095
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   360
            Width           =   330
         End
         Begin VB.Image imgCalendario2 
            Height          =   360
            Left            =   14685
            Picture         =   "frmRH_Funcionarios.frx":11518
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   930
            Width           =   330
         End
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
         Height          =   825
         Left            =   -74945
         TabIndex        =   133
         Top             =   1305
         Width           =   15200
         Begin VB.CommandButton Cmd_localizar_caminho_ates 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14010
            Picture         =   "frmRH_Funcionarios.frx":1199B
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Localizar atestado."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho_ates 
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
            Left            =   8310
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   60
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do atestado."
            Top             =   390
            Width           =   5685
         End
         Begin VB.CommandButton Cmd_limpar_caminho_ates 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14340
            Picture         =   "frmRH_Funcionarios.frx":11A9D
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Limpar caminho."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_ates 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmRH_Funcionarios.frx":11BDB
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Visualizar arquivo."
            Top             =   390
            Width           =   315
         End
         Begin VB.ComboBox Cmb_tipo_ates 
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
            ItemData        =   "frmRH_Funcionarios.frx":1219D
            Left            =   5760
            List            =   "frmRH_Funcionarios.frx":121A7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   58
            ToolTipText     =   "Tipo."
            Top             =   390
            Width           =   1125
         End
         Begin VB.TextBox Txt_ID_ates 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1410
            TabIndex        =   134
            Text            =   "0"
            Top             =   390
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox Txt_data_ates 
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
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   975
         End
         Begin VB.TextBox Txt_responsavel_ates 
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
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3160
         End
         Begin MSComCtl2.DTPicker Cmb_emissao_ates 
            Height          =   315
            Left            =   4350
            TabIndex        =   57
            ToolTipText     =   "Data de emissão."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   198705153
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker Cmb_validade_ates 
            Height          =   315
            Left            =   6900
            TabIndex        =   59
            ToolTipText     =   "Data de validade."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   198705153
            CurrentDate     =   39057
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. validade"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7162
            TabIndex        =   140
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho do atestado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10387
            TabIndex        =   139
            Top             =   180
            Width           =   1530
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6135
            TabIndex        =   138
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   495
            TabIndex        =   137
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2295
            TabIndex        =   136
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. emissão"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4627
            TabIndex        =   135
            Top             =   180
            Width           =   840
         End
      End
      Begin VB.Frame Frame16 
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
         Height          =   825
         Left            =   -74945
         TabIndex        =   124
         Top             =   1305
         Width           =   15200
         Begin VB.TextBox txtIdObs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            TabIndex        =   131
            Text            =   "0"
            Top             =   360
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtObs 
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
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   66
            ToolTipText     =   "Observações."
            Top             =   390
            Width           =   13305
         End
         Begin MSComCtl2.DTPicker txtData_Obs 
            Height          =   315
            Left            =   180
            TabIndex        =   65
            ToolTipText     =   "Data."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   198705153
            CurrentDate     =   39057
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7785
            TabIndex        =   126
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   690
            TabIndex        =   125
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.TextBox txtid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6150
         TabIndex        =   118
         Text            =   "0"
         Top             =   1710
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Frame Frame11 
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
         Height          =   1425
         Left            =   -74945
         TabIndex        =   93
         Top             =   1305
         Width           =   15200
         Begin VB.TextBox txtIdCurso 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   510
            TabIndex        =   127
            Text            =   "0"
            Top             =   390
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14400
            Picture         =   "frmRH_Funcionarios.frx":121B8
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Limpar caminho."
            Top             =   960
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14730
            Picture         =   "frmRH_Funcionarios.frx":122F6
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Visualizar arquivo."
            Top             =   960
            Width           =   315
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   7410
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14070
            Picture         =   "frmRH_Funcionarios.frx":128B8
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Localizar certificado de treinamento."
            Top             =   960
            Width           =   315
         End
         Begin VB.TextBox Txt_caminho 
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
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do certificado de treinamento."
            Top             =   960
            Width           =   13875
         End
         Begin VB.TextBox txtNivel 
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
            Left            =   12060
            MaxLength       =   10
            TabIndex        =   35
            ToolTipText     =   "Nível de qualificação."
            Top             =   390
            Width           =   1725
         End
         Begin VB.CommandButton cmdCurso 
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
            Left            =   4440
            Picture         =   "frmRH_Funcionarios.frx":129BA
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Localizar cursos."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtInstituicao 
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
            MaxLength       =   255
            TabIndex        =   34
            ToolTipText     =   "Instituição educacional."
            Top             =   390
            Width           =   5595
         End
         Begin VB.TextBox txtduracaocurso 
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
            Left            =   13800
            MaxLength       =   30
            TabIndex        =   36
            ToolTipText     =   "Duração."
            Top             =   390
            Width           =   1245
         End
         Begin VB.TextBox txtCurso 
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
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Curso."
            Top             =   390
            Width           =   4245
         End
         Begin MSMask.MaskEdBox txtDtCurso 
            Height          =   315
            Left            =   4890
            TabIndex        =   33
            ToolTipText     =   "Data."
            Top             =   390
            Width           =   1125
            _ExtentX        =   1984
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
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho do certificado de treinamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5730
            TabIndex        =   117
            Top             =   750
            Width           =   2775
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nível de qualificação*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12150
            TabIndex        =   111
            Top             =   180
            Width           =   1545
         End
         Begin VB.Image imgCalendario3 
            Height          =   360
            Left            =   6030
            Picture         =   "frmRH_Funcionarios.frx":12ABC
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instituição educacional*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8385
            TabIndex        =   97
            Top             =   180
            Width           =   1725
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5235
            TabIndex        =   96
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duração"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   14122
            TabIndex        =   95
            Top             =   180
            Width           =   600
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Curso*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2047
            TabIndex        =   94
            Top             =   180
            Width           =   510
         End
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9960
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmRH_Funcionarios.frx":12F3F
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   -65970
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmRH_Funcionarios.frx":1A293
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -64410
         Top             =   510
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmRH_Funcionarios.frx":1F677
         Count           =   1
      End
      Begin VB.Frame Frame15 
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
         Height          =   1425
         Left            =   -74945
         TabIndex        =   108
         Top             =   1305
         Width           =   15200
         Begin VB.TextBox txtIdSindicato 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   130
            Text            =   "0"
            Top             =   750
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox Txt_obs_sindicato 
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
            Height          =   885
            Left            =   1590
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            ToolTipText     =   "Observações."
            Top             =   390
            Width           =   13395
         End
         Begin VB.TextBox txtValor_Sindicato 
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
            MaxLength       =   30
            TabIndex        =   52
            ToolTipText     =   "Valor."
            Top             =   960
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker txtData_Sindicato 
            Height          =   315
            Left            =   180
            TabIndex        =   51
            ToolTipText     =   "Data."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   173080577
            CurrentDate     =   39057
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7800
            TabIndex        =   141
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   705
            TabIndex        =   110
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   652
            TabIndex        =   109
            Top             =   750
            Width           =   450
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Período aquisitivo"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   -71535
         TabIndex        =   105
         Top             =   1305
         Width           =   11805
         Begin VB.TextBox txtIdFerias 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   510
            TabIndex        =   129
            Text            =   "0"
            Top             =   300
            Visible         =   0   'False
            Width           =   585
         End
         Begin MSComCtl2.DTPicker txtInicio_Periodo_Ferias 
            Height          =   315
            Left            =   2700
            TabIndex        =   48
            ToolTipText     =   "Data início."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   173080577
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtFim_Periodo_Ferias 
            Height          =   315
            Left            =   4290
            TabIndex        =   49
            ToolTipText     =   "Data final."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   173080577
            CurrentDate     =   39057
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3300
            TabIndex        =   107
            Top             =   180
            Width           =   195
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4830
            TabIndex        =   106
            Top             =   180
            Width           =   255
         End
      End
      Begin VB.Frame Frame13 
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
         Height          =   825
         Left            =   -74945
         TabIndex        =   102
         Top             =   1305
         Width           =   3405
         Begin MSComCtl2.DTPicker txtData_Inicio_Ferias 
            Height          =   315
            Left            =   180
            TabIndex        =   46
            ToolTipText     =   "Data início."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   173080577
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtData_Fim_Ferias 
            Height          =   315
            Left            =   1785
            TabIndex        =   47
            ToolTipText     =   "Data final."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   173080577
            CurrentDate     =   39057
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   780
            TabIndex        =   104
            Top             =   180
            Width           =   195
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2355
            TabIndex        =   103
            Top             =   180
            Width           =   255
         End
      End
      Begin VB.Frame Frame12 
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
         Height          =   1425
         Left            =   -74945
         TabIndex        =   98
         Top             =   1305
         Width           =   15200
         Begin VB.TextBox txtIdAumento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3420
            TabIndex        =   128
            Text            =   "0"
            Top             =   630
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtObsAumentos 
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
            Height          =   885
            Left            =   1320
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            ToolTipText     =   "Observações."
            Top             =   390
            Width           =   13665
         End
         Begin VB.TextBox txtValorAumento 
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
            TabIndex        =   43
            ToolTipText     =   "Valor."
            Top             =   960
            Width           =   1125
         End
         Begin MSMask.MaskEdBox txtMesAno 
            Height          =   315
            Left            =   180
            TabIndex        =   42
            ToolTipText     =   "Mês/ano."
            Top             =   390
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
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
            Mask            =   "##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7680
            TabIndex        =   101
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   517
            TabIndex        =   100
            Top             =   750
            Width           =   450
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês/Ano*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   382
            TabIndex        =   99
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   75
         Top             =   1200
         Width           =   11820
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   79
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do contato:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   405
            TabIndex        =   78
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ramal:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   1200
            TabIndex        =   77
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   1215
            TabIndex        =   76
            Top             =   1478
            Width           =   480
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   113
         Top             =   330
         Width           =   15200
         _ExtentX        =   26802
         _ExtentY        =   1720
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
         ButtonCaption8  =   "Filtrar todos"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Carregar/Limpar lista de funcionários cadastrados (F7)"
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
         ButtonWidth8    =   77
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Validação"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Validação (F9)"
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
         ButtonLeft9     =   432
         ButtonTop9      =   2
         ButtonWidth9    =   53
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
         ButtonLeft10    =   487
         ButtonTop10     =   4
         ButtonWidth10   =   2
         ButtonHeight10  =   54
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
         ButtonLeft11    =   491
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
         ButtonLeft12    =   534
         ButtonTop12     =   2
         ButtonWidth12   =   30
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   566
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74945
         TabIndex        =   115
         Top             =   330
         Width           =   15200
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   8
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
         ButtonLeft2     =   48
         ButtonTop2      =   2
         ButtonWidth2    =   60
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   110
         ButtonTop3      =   2
         ButtonWidth3    =   55
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   167
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonAlignment5=   2
         ButtonType5     =   1
         ButtonStyle5    =   -1
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   -1
         ButtonLeft5     =   224
         ButtonTop5      =   4
         ButtonWidth5    =   2
         ButtonHeight5   =   54
         ButtonCaption6  =   "Ajuda"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Ajuda (F1)"
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
         ButtonLeft6     =   228
         ButtonTop6      =   2
         ButtonWidth6    =   41
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Sair"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Sair (Esc)"
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
         ButtonLeft7     =   271
         ButtonTop7      =   2
         ButtonWidth7    =   30
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   5
         ButtonLeft8     =   303
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
      End
   End
End
Attribute VB_Name = "frmRH_Funcionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoUsuario                    As String 'OK
Dim NumeroUsuario                    As Integer 'OK
Dim Novo_Funcionario                 As Boolean 'OK
Dim Novo_Funcionario2                As Boolean 'OK
Dim Novo_Funcionario3                As Boolean 'OK
Dim Novo_Funcionario4                As Boolean 'OK
Dim Novo_Funcionario5                As Boolean 'OK
Dim Novo_Funcionario6                As Boolean 'OK
Dim Novo_Funcionario7                As Boolean 'OK
Dim Novo_Funcionario8                As Boolean 'OK
Dim TBLISTA_RH_Funcionarios          As ADODB.Recordset 'OK
Public StrSql_Localizar_Funcionarios As String 'OK
Public FormulaRel_Funcionarios       As String 'OK
Public Aniversario                   As Boolean

'Corrige formulario
Dim Top_Lista As Long
Dim Height_Lista As Long

Private Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=xgs8lepT1LI")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

Top_Lista = Lista.Top
Height_Lista = Lista.Height

If SSTab1.Tab = 1 Then
    With Lista
        .Visible = True
        .Top = Top_Lista - (Frame2.Height - Frame3.Height)
        .Height = Height_Lista + (Frame2.Height - Frame3.Height)
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregarTodos()
On Error GoTo tratar_erro

StrSql_Localizar_Funcionarios = "Select * from Funcionarios order by nome"
Aniversario = False
ProcAtualizalista (1)

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
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(9) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(9) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_ates_Click()
On Error GoTo tratar_erro

Txt_caminho_ates = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

txt_Caminho = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_caminho_ates_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho_ates = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_doc_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho_doc = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_foto_Click()
On Error GoTo tratar_erro

fotopadrao = ""
CommonDialog1.Filter = "Arquivos jpg (*.jpg) | *.jpg| Arquivos bmp (*.bmp) | *.bmp"

'Diretorio onde estão as imagens
CommonDialog1.InitDir = App.Path
CommonDialog1.DefaultExt = "*.*"
CommonDialog1.ShowOpen
fotopadrao = CommonDialog1.filename

If fotopadrao <> "" Then Foto.Picture = LoadPicture(fotopadrao) Else Foto.Picture = LoadPicture("")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_ates_Click()
On Error GoTo tratar_erro

If Txt_caminho_ates <> "" Then ProcAbrirArquivo Txt_caminho_ates

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_doc_Click()
On Error GoTo tratar_erro

If Txt_caminho_doc <> "" Then ProcAbrirArquivo Txt_caminho_doc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFornecedor_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCurso_Click()
On Error GoTo tratar_erro

frmRH_Funcionarios_descricao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirCurso()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaCursos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) curso(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from funcionarios_cursos where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            Evento = "Excluir curso"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
            Documento1 = "Curso: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) curso(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Curso(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCurso
    ProcAtualizalistaCursos
    Frame11.Enabled = False
    Novo_Funcionario2 = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirAumento()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaAumentos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) aumento(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from funcionarios_aumentos where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            Evento = "Excluir aumento"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
            Documento1 = "Mês/Ano: " & .ListItems(InitFor).SubItems(1) & " - Valor: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) aumento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Aumento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposAumentos
    ProcAtualizalistaAumentos
    Frame12.Enabled = False
    Novo_Funcionario3 = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirFerias()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaFerias
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir estas férias?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from funcionarios_ferias where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            Evento = "Excluir férias"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
            Documento1 = "De: " & .ListItems(InitFor).SubItems(1) & " - Até: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe as férias antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Férias excluídas com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposFerias
    ProcAtualizalistaFerias
    Frame13.Enabled = False
    Frame14.Enabled = False
    Novo_Funcionario4 = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirSindicato()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaSindicato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) sindicato(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from funcionarios_sindicato where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            Evento = "Excluir sindicato"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
            Documento1 = "Data: " & .ListItems(InitFor).SubItems(1) & " - Valor: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) sindicato(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Sindicato(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposSindicato
    ProcAtualizalistaSindicato
    Frame15.Enabled = False
    Novo_Funcionario5 = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirAtes()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaAtestados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) atestado(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Funcionarios_atestados where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            Evento = "Excluir atestado"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
            Documento1 = "Data de emissão: " & .ListItems(InitFor).ListSubItems(3) & " - Tipo: " & .ListItems(InitFor).ListSubItems(4) & " - Data de validade: " & .ListItems(InitFor).ListSubItems(5)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) atestado(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Atestado(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposAtes
    ProcAtualizalistaAtes
    Frame1.Enabled = False
    Novo_Funcionario6 = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirObs()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaObs
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) observação(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from funcionarios_obs where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            Evento = "Excluir observação"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
            Documento1 = "Data: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) observação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Observação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposObs
    ProcAtualizalistaObs
    Frame15.Enabled = False
    Novo_Funcionario7 = False
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_doc()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) documento(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            '==================================
            Modulo = "RH/Cadastro de Funcionários"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & txtdesenhoproduto
            Documento1 = "Caminho: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
            Conexao.Execute "DELETE from Funcionarios_documentos where ID = " & .ListItems(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) documento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Documento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proclimpacampos_doc
    ProcAtualizaLista_Doc
    Novo_Funcionario8 = False
    Frame4.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
frmRH_funcionarios_MenuImpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoCurso()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novos", txtDtValidacao, "funcionário", "cursos", True) = False Then Exit Sub
ProcLimpaCamposCurso
Frame11.Enabled = True
Novo_Funcionario2 = True
txtCurso.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoAumento()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "funcionário", "aumento", True) = False Then Exit Sub
ProcLimpaCamposAumentos
Frame12.Enabled = True
Novo_Funcionario3 = True
txtMesAno.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoFerias()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "funcionário", "férias", True) = False Then Exit Sub
ProcLimpaCamposFerias
Frame13.Enabled = True
Frame14.Enabled = True
Novo_Funcionario4 = True
txtData_Inicio_Ferias.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoSindicato()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "funcionário", "sindicato", True) = False Then Exit Sub
ProcLimpaCamposSindicato
Frame15.Enabled = True
Novo_Funcionario5 = True
txtData_Sindicato.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoAtes()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "funcionário", "atestado", True) = False Then Exit Sub
ProcLimpaCamposAtes
Frame1.Enabled = True
Novo_Funcionario6 = True
Cmb_emissao_ates.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoObs()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar nova", txtDtValidacao, "funcionário", "observação", True) = False Then Exit Sub
ProcLimpaCamposObs
Frame16.Enabled = True
Novo_Funcionario7 = True
txtData_Obs.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_doc()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "funcionário", "documento", True) = False Then Exit Sub
Proclimpacampos_doc
Novo_Funcionario8 = True
Frame4.Enabled = True
Cmd_localizar_doc_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposCurso()
On Error GoTo tratar_erro

txtIdCurso = 0
txtCurso = ""
txtDtCurso = "__/__/____"
txtInstituicao = ""
txtduracaocurso = ""
txtNivel = ""
txt_Caminho = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposAumentos()
On Error GoTo tratar_erro

txtIdAumento = 0
txtMesAno = "__/____"
txtValorAumento = ""
txtObsAumentos = ""
CodigoLista3 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposFerias()
On Error GoTo tratar_erro

txtIdFerias = 0
txtData_Inicio_Ferias.Value = Date
txtData_Fim_Ferias.Value = Date
txtInicio_Periodo_Ferias.Value = Date
txtFim_Periodo_Ferias.Value = Date
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposSindicato()
On Error GoTo tratar_erro

txtIdSindicato = 0
txtData_Sindicato.Value = Date
txtValor_Sindicato = ""
Txt_obs_sindicato = ""
CodigoLista5 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposAtes()
On Error GoTo tratar_erro

Txt_ID_ates = 0
Txt_data_ates = Format(Date, "dd/mm/yy")
Txt_responsavel_ates = pubUsuario
Cmb_emissao_ates.Value = Date
Cmb_tipo_ates = "Médico"
Cmb_validade_ates.Value = Date
Txt_caminho_ates = ""
CodigoLista6 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposObs()
On Error GoTo tratar_erro

txtIdObs = 0
txtData_Obs.Value = Date
txtObs = ""
CodigoLista7 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proclimpacampos_doc()
On Error GoTo tratar_erro

txtID_doc = 0
txtData_doc = Format(Date, "dd/mm/yy")
txtResponsavel_doc = pubUsuario
Txt_caminho_doc = ""
Txt_obs_doc = ""
CodigoLista8 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar1()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "dados do funcionário", True) = False Then Exit Sub
If cmbSituacao = "Afastado" Then
    If txtdata_desligado = "__/__/___" Then
        USMsgBox "Informe a data de afastamento antes de salvar.", vbInformation, "CAPRIND v5.0"
        txtdata_desligado.SetFocus
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Funcionarios where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!dependente = IIf(txtDependente = "", Null, txtDependente)
TBGravar!Funcao = txtFuncao
TBGravar!Tipo = txttipo
If txtData_adm <> "__/__/____" Then TBGravar!Admissao = txtData_adm Else TBGravar!Admissao = Null
TBGravar!Setor = txtSetor
TBGravar!divisao = txtDivisao
TBGravar!Horario_inicio = Format(txtHorario_inicio, "hh:mm")
TBGravar!Horario_fim = Format(txtHorario_fim, "hh:mm")
TBGravar!Turno = txtTurno
TBGravar!IDFornecedor = IIf(txtIDfornecedor = "", Null, txtIDfornecedor)
TBGravar!conta_corrente = txtConta_corrente
TBGravar!Salario = IIf(txtSalario = "", 0, txtSalario)
TBGravar!situacao = cmbSituacao
If txtdata_desligado <> "__/__/____" Then TBGravar!data_desligado = txtdata_desligado Else TBGravar!data_desligado = Null
TBGravar.Update
TBGravar.Close
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If
1:
    USMsgBox ("Dados do funcionário cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "RH/Funcionários"
    Evento = "Dados do funcionário"
    ID_documento = txtId
    Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
    Documento1 = ""
    ProcGravaEvento
    '==================================

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCurso()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame11.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "cursos", True) = False Then Exit Sub
Acao = "salvar"
If txtCurso = "" Then
    NomeCampo = "o curso"
    ProcVerificaAcao
    frmRH_Funcionarios_descricao.Show 1
    Exit Sub
End If
If txtDtCurso = "__/__/____" Then
    NomeCampo = "a data"
    ProcVerificaAcao
    txtDtCurso.SetFocus
    Exit Sub
End If
If txtInstituicao = "" Then
    NomeCampo = "a instituição educacional"
    ProcVerificaAcao
    txtInstituicao.SetFocus
    Exit Sub
End If
If txtNivel = "" Then
    NomeCampo = "o nível de qualificação"
    ProcVerificaAcao
    txtNivel.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from funcionarios_cursos where id = " & txtIdCurso, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosCurso
TBGravar.Update
txtIdCurso = TBGravar!ID
TBGravar.Close
ProcAtualizalistaCursos
If Novo_Funcionario2 = True Then
    USMsgBox ("Novo curso cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo curso"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar curso"
    If CodigoLista2 <> 0 And ListaCursos.ListItems.Count <> 0 Then
        ListaCursos.SelectedItem = ListaCursos.ListItems(CodigoLista2)
        ListaCursos.SetFocus
    End If
End If
'==================================
Modulo = "RH/Cadastro de funcionários"
ID_documento = txtIdCurso
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "Curso: " & txtCurso
ProcGravaEvento
'==================================
Novo_Funcionario2 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarAumento()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame12.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "aumento", True) = False Then Exit Sub
Acao = "salvar"
If txtMesAno = "__/____" Then
    NomeCampo = "o mês/ano"
    ProcVerificaAcao
    txtMesAno.SetFocus
    Exit Sub
End If
If txtValorAumento = "" Or txtValorAumento = "0,00" Then
    NomeCampo = "o valor do aumento"
    ProcVerificaAcao
    txtValorAumento.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from funcionarios_aumentos where id = " & txtIdAumento, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosAumentos
TBGravar.Update
txtIdAumento = TBGravar!ID
TBGravar.Close
ProcAtualizalistaAumentos
If Novo_Funcionario3 = True Then
    USMsgBox ("Novo aumento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo aumento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar aumento"
    If CodigoLista3 <> 0 And ListaAumentos.ListItems.Count <> 0 Then
        ListaAumentos.SelectedItem = ListaAumentos.ListItems(CodigoLista3)
        ListaAumentos.SetFocus
    End If
End If
Novo_Funcionario3 = False
'==================================
Modulo = "RH/Cadastro de Funcionários"
ID_documento = txtIdAumento
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "Mês/Ano: " & txtMesAno & " - Valor: " & txtValorAumento
ProcGravaEvento
'==================================
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarFerias()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame13.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "férias", True) = False Then Exit Sub
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from funcionarios_ferias where id = " & txtIdFerias, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosFerias
TBGravar.Update
txtIdFerias = TBGravar!ID
TBGravar.Close
ProcAtualizalistaFerias
If Novo_Funcionario4 = True Then
    USMsgBox ("Novas férias cadastradas com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova férias"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar férias"
    If CodigoLista4 <> 0 And ListaFerias.ListItems.Count <> 0 Then
        ListaFerias.SelectedItem = ListaFerias.ListItems(CodigoLista4)
        ListaFerias.SetFocus
    End If
End If
Novo_Funcionario4 = False
'==================================
Modulo = "RH/Cadastro de Funcionários"
ID_documento = txtIdFerias
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "De: " & txtData_Inicio_Ferias & " - Até: " & txtData_Fim_Ferias
ProcGravaEvento
'==================================
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarSindicato()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame15.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "sindicato", True) = False Then Exit Sub
If txtValor_Sindicato = "" Or txtValor_Sindicato = "0,00" Then
    USMsgBox ("Informe o valor do sindicato antes de salvar."), vbInformation, "CAPRIND v5.0"
    txtValor_Sindicato.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from funcionarios_sindicato where id = " & txtIdSindicato, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosSindicato
TBGravar.Update
txtIdSindicato = TBGravar!ID
TBGravar.Close
ProcAtualizalistaSindicato
If Novo_Funcionario5 = True Then
    USMsgBox ("Novo sindicato cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo sindicato"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar sindicato"
    If CodigoLista5 <> 0 And ListaSindicato.ListItems.Count <> 0 Then
        ListaSindicato.SelectedItem = ListaSindicato.ListItems(CodigoLista5)
        ListaSindicato.SetFocus
    End If
End If
Novo_Funcionario5 = False
'==================================
Modulo = "RH/Cadastro de Funcionários"
ID_documento = txtIdSindicato
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "Data: " & txtData_Sindicato & " - Valor: " & txtValor_Sindicato
ProcGravaEvento
'==================================
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarAtes()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "atestado", True) = False Then Exit Sub
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Funcionarios_atestados where ID = " & Txt_ID_ates, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosAtes
TBGravar.Update
Txt_ID_ates = TBGravar!ID
TBGravar.Close
ProcAtualizalistaAtes
If Novo_Funcionario6 = True Then
    USMsgBox ("Novo atestado cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo atestado"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar atestado"
    If CodigoLista6 <> 0 And ListaAtestados.ListItems.Count <> 0 Then
        ListaAtestados.SelectedItem = ListaAtestados.ListItems(CodigoLista6)
        ListaAtestados.SetFocus
    End If
End If
'==================================
Modulo = "RH/Cadastro de Funcionários"
ID_documento = Txt_ID_ates
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "Data de emissão: " & Format(Cmb_emissao_ates, "dd/mm/yy") & " - Tipo: " & Cmb_tipo_ates & " - Data de validade: " & Format(Cmb_validade_ates, "dd/mm/yy")
ProcGravaEvento
'==================================
Novo_Funcionario6 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarObs()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame16.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "observação", True) = False Then Exit Sub
If txtObs = "" Then
    USMsgBox ("Informe a observação antes de salvar."), vbInformation, "CAPRIND v5.0"
    txtObs.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from funcionarios_obs where id = " & txtIdObs, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosObs
TBGravar.Update
txtIdObs = TBGravar!ID
TBGravar.Close
ProcAtualizalistaObs
If Novo_Funcionario7 = True Then
    USMsgBox ("Nova observação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova observação"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar observação"
    If CodigoLista7 <> 0 And ListaObs.ListItems.Count <> 0 Then
        ListaObs.SelectedItem = ListaObs.ListItems(CodigoLista7)
        ListaObs.SetFocus
    End If
End If
'==================================
Modulo = "RH/Cadastro de Funcionários"
ID_documento = txtIdObs
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "Data: " & txtData_Obs
ProcGravaEvento
'==================================
Novo_Funcionario7 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_doc()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "funcionário", "documento", True) = False Then Exit Sub
Acao = "salvar"
If Txt_caminho_doc = "" Then
    NomeCampo = "o caminho"
    ProcVerificaAcao
    Cmd_localizar_doc.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Funcionarios_documentos where ID = " & txtID_doc, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviadadosDoc
TBGravar.Update
txtID_doc = TBGravar!ID
TBGravar.Close
ProcAtualizaLista_Doc
If Novo_Funcionario8 = True Then
    USMsgBox ("Novo documento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo documento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar documento"
    If CodigoLista8 <> 0 And Lista_doc.ListItems.Count <> 0 Then
        Lista_doc.SelectedItem = Lista_doc.ListItems(CodigoLista1)
        Lista_doc.SetFocus
    End If
End If
'==================================
Modulo = "RH/Cadastro de Funcionários"
ID_documento = txtID_doc
Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
Documento1 = "Caminho: " & txt_Caminho
ProcGravaEvento
'==================================
Novo_Funcionario8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosCurso()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Descricao = txtCurso
If txtDtCurso <> "__/__/____" Then TBGravar!Data = txtDtCurso Else TBGravar!Data = Null
If txtduracaocurso <> "" Then TBGravar!duracao = txtduracaocurso Else TBGravar!duracao = Null
TBGravar!Instituicao = txtInstituicao
TBGravar!Nivel = txtNivel
TBGravar!caminho = txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosAumentos()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Data = Format(txtMesAno, "mm/yyyy")
TBGravar!valor = txtValorAumento
TBGravar!Obs = txtObsAumentos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosFerias()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Data_inicio = txtData_Inicio_Ferias.Value
TBGravar!Data_fim = txtData_Fim_Ferias.Value
TBGravar!inicio_periodo = txtInicio_Periodo_Ferias.Value
TBGravar!fim_periodo = txtFim_Periodo_Ferias.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosSindicato()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Data = txtData_Sindicato.Value
TBGravar!valor = IIf(txtValor_Sindicato = "", Null, txtValor_Sindicato)
TBGravar!Obs = Txt_obs_sindicato

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosAtes()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Data = Txt_data_ates
TBGravar!Responsavel = Txt_responsavel_ates
TBGravar!Data_emissao = Cmb_emissao_ates
TBGravar!Tipo = Left(Cmb_tipo_ates, 1)
TBGravar!Data_validade = Cmb_validade_ates
TBGravar!caminho = Txt_caminho_ates

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosObs()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Data = txtData_Obs.Value
TBGravar!Obs = txtObs

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosDoc()
On Error GoTo tratar_erro

TBGravar!ID_funcionario = txtId
TBGravar!Data = IIf(txtData_doc = "", Date, txtData_doc)
TBGravar!Responsavel = IIf(txtResponsavel_doc = "", pubUsuario, txtResponsavel_doc)
TBGravar!caminho = Txt_caminho_doc
TBGravar!Obs = Trim(Txt_obs_doc)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizalistaCursos()
On Error GoTo tratar_erro

ListaCursos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from funcionarios_cursos where id_funcionario = " & txtId & " order by descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaCursos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = Trim(TBLISTA!Descricao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yyyy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!duracao), "", TBLISTA!duracao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Instituicao), "", TBLISTA!Instituicao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Nivel), "", TBLISTA!Nivel)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizalistaAumentos()
On Error GoTo tratar_erro

ListaAumentos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from funcionarios_aumentos where id_funcionario = " & txtId & " order by data desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaAumentos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "mm/yyyy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Obs), "", Trim(TBLISTA!Obs))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizalistaFerias()
On Error GoTo tratar_erro

ListaFerias.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from funcionarios_ferias where id_funcionario = " & txtId & " order by data_inicio desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaFerias.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data_inicio), "", Format(TBLISTA!Data_inicio, "dd/mm/yyyy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data_fim), "", Format(TBLISTA!Data_fim, "dd/mm/yyyy"))
            .Item(.Count).SubItems(3) = DateDiff("d", TBLISTA!Data_inicio, TBLISTA!Data_fim) + 1
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!inicio_periodo), "", Format(TBLISTA!inicio_periodo, "dd/mm/yyyy"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!fim_periodo), "", Format(TBLISTA!fim_periodo, "dd/mm/yyyy"))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizalistaSindicato()
On Error GoTo tratar_erro

ListaSindicato.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from funcionarios_sindicato where id_funcionario = " & txtId & " order by data desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaSindicato.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yyyy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizalistaAtes()
On Error GoTo tratar_erro

ListaAtestados.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Funcionarios_atestados where ID_funcionario = " & txtId & " order by Data_emissao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaAtestados.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Data_emissao), "", Format(TBLISTA!Data_emissao, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Tipo), "", IIf(TBLISTA!Tipo = "M", "Médico", "ASO"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data_validade), "", Format(TBLISTA!Data_validade, "dd/mm/yy"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizalistaObs()
On Error GoTo tratar_erro

ListaObs.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from funcionarios_obs where id_funcionario = " & txtId & " order by data desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaObs.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Obs), "", Trim(TBLISTA!Obs))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaLista_Doc()
On Error GoTo tratar_erro

Lista_doc.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ID, Caminho from Funcionarios_documentos where ID_funcionario = " & txtId & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_doc.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbSituacao_Click()
On Error GoTo tratar_erro

With txtdata_desligado
    If cmbSituacao = "Afastado" Or cmbSituacao = "Demitido" Then
        .Enabled = True
        If cmbSituacao = "Afastado" Then
            Label26.Caption = "Dt. afastado"
            .ToolTipText = "Data do afastamento."
        Else
            Label26.Caption = "Dt. demitido"
            .ToolTipText = "Data da demisão."
        End If
    Else
        Label26.Caption = "Dt. afastado"
        .ToolTipText = "Data do afastamento."
        .Enabled = False
        .Text = "__/__/____"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDependentes_Click()
On Error GoTo tratar_erro

frmRH_Funcionarios_dependentes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDivisao_Click()
On Error GoTo tratar_erro

frmRH_Funcionarios_divisao.Show 1

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
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) funcionário(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute ("DELETE from Funcionarios where id = " & .ListItems(InitFor))
            Conexao.Execute ("DELETE from Funcionarios_aumentos where id_funcionario = " & .ListItems(InitFor))
            Conexao.Execute ("DELETE from Funcionarios_cursos where id_funcionario = " & .ListItems(InitFor))
            Conexao.Execute ("DELETE from Funcionarios_ferias where id_funcionario = " & .ListItems(InitFor))
            Conexao.Execute ("DELETE from Funcionarios_obs where id_funcionario = " & .ListItems(InitFor))
            Conexao.Execute ("DELETE from Funcionarios_sindicato where id_funcionario = " & .ListItems(InitFor))
            '==================================
            Modulo = "RH/Funcionários"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Código: " & .ListItems(InitFor).SubItems(1) & " - Funcionário: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) funcionário(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Funcionário(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (1)
    ProcLimpaCampos
    Frame2.Enabled = False
    Frame3.Enabled = False
    Novo_Funcionario = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdfuncao_Click()
On Error GoTo tratar_erro

frmRH_Funcionarios_funcao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmRH_funcionarios_localizar.Show 1

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
Frame2.Enabled = True
Frame3.Enabled = True
Novo_Funcionario = True

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select isnumeric(Codigo) as codigo from Funcionarios", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.Find ("Codigo = 0")
    If TBAbrir.EOF = True Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select CONVERT(Numeric(14), Codigo) as Codigo from Funcionarios order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            txtCodigo = TBAbrir!CODIGO + 1
        Else
            txtCodigo = "001"
        End If
        TBAbrir.Close
    End If
End If
txtCodigo.SetFocus

ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Funcionario = True Then
    If USMsgBox("O funcionário ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Funcionario = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario2 = True Then
    If USMsgBox("O curso ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 2
        ProcSalvarCurso
        If Novo_Funcionario2 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario3 = True Then
    If USMsgBox("O aumento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 3
        ProcSalvarAumento
        If Novo_Funcionario3 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario4 = True Then
    If USMsgBox("As férias ainda não foram salvas, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 4
        ProcSalvarFerias
        If Novo_Funcionario4 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario5 = True Then
    If USMsgBox("O sindicato ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 5
        ProcSalvarSindicato
        If Novo_Funcionario5 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario6 = True Then
    If USMsgBox("O atestado ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 6
        ProcSalvarAtes
        If Novo_Funcionario6 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario7 = True Then
    If USMsgBox("A observação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 7
        ProcSalvarObs
        If Novo_Funcionario7 = True Then Exit Sub Else Unload Me
    End If
End If
If Novo_Funcionario8 = True Then
    If USMsgBox("O documentos ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        SSTab1.Tab = 8
        procSalvar_doc
        If Novo_Funcionario8 = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Funcionario = False
Novo_Funcionario2 = False
Novo_Funcionario3 = False
Novo_Funcionario4 = False
Novo_Funcionario5 = False
Novo_Funcionario6 = False
Novo_Funcionario7 = False
Novo_Funcionario8 = False
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
If Frame2.Enabled = False Then
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
If txtCodigo = "" Then
    NomeCampo = "o código"
    ProcVerificaAcao
    txtCodigo.SetFocus
    Exit Sub
End If
'Verifica se o código está sendo utilizado
Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select * from funcionarios where codigo = '" & txtCodigo & "' and id <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = False Then
    USMsgBox ("Esse código está sendo utilizado, favor alterar."), vbInformation, "CAPRIND v5.0"
    txtCodigo = ""
    txtCodigo.SetFocus
    Exit Sub
End If
TBCodigoDesc.Close
If txtNome = "" Then
    USMsgBox ("Informe o nome do funcionário antes de salvar."), vbInformation, "CAPRIND v5.0"
    txtNome.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Funcionarios where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "o mesmo", "funcionário", True) = False Then Exit Sub
    If txtNome <> TBGravar!Nome Then
        Conexao.Execute "Update CFI Set Funcionario = '" & txtNome & "' where Funcionario = '" & TBGravar!Nome & "'"
        Conexao.Execute "Update tbl_ContasPagar Set txt_Fornecedor = '" & txtNome & "' where int_codforn = " & IIf(txtId = "", 0, txtId) & " and Tipo = 'FU'"
    End If
Else
    TBGravar.AddNew
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
If Novo_Funcionario = True Then
    USMsgBox ("Novo funcionário cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    ProcAtualizalista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "RH/Funcionários"
    ID_documento = txtId
    Documento = "Código: " & txtCodigo & " - Funcionário: " & txtNome
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_Funcionario = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_RH_Funcionarios.AbsolutePage <> 2 Then
    If TBLISTA_RH_Funcionarios.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_RH_Funcionarios.PageCount - 1)
    Else
        TBLISTA_RH_Funcionarios.AbsolutePage = TBLISTA_RH_Funcionarios.AbsolutePage - 2
        ProcExibePagina (TBLISTA_RH_Funcionarios.AbsolutePage)
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
    TBLISTA_RH_Funcionarios.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_RH_Funcionarios.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_RH_Funcionarios.AbsolutePage = 1
ProcExibePagina (TBLISTA_RH_Funcionarios.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_RH_Funcionarios.AbsolutePage <> -3 Then
    If TBLISTA_RH_Funcionarios.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_RH_Funcionarios.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_RH_Funcionarios.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_RH_Funcionarios.AbsolutePage = TBLISTA_RH_Funcionarios.PageCount
ProcExibePagina (TBLISTA_RH_Funcionarios.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSetor_Click()
On Error GoTo tratar_erro

CadMaquinas = False
Funcionario = True
Usuarios = False
Estoque_Local_Armazenamento = False
frmUsuarios_Setor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdTurno_Click()
On Error GoTo tratar_erro

frmRH_Funcionarios_turno.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

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
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCarregarTodos
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyF3: procSalvar1
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoCurso
            Case vbKeyF3: ProcSalvarCurso
            Case vbKeyF4: ProcExcluirCurso
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoAumento
            Case vbKeyF3: ProcSalvarAumento
            Case vbKeyF4: ProcExcluirAumento
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoFerias
            Case vbKeyF3: ProcSalvarFerias
            Case vbKeyF4: ProcExcluirFerias
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 5:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoSindicato
            Case vbKeyF3: ProcSalvarSindicato
            Case vbKeyF4: ProcExcluirSindicato
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 6:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoAtes
            Case vbKeyF3: ProcSalvarAtes
            Case vbKeyF4: ProcExcluirAtes
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 7:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoObs
            Case vbKeyF3: ProcSalvarObs
            Case vbKeyF4: ProcExcluirObs
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 8:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_doc
            Case vbKeyF3: procSalvar_doc
            Case vbKeyF4: procExcluir_doc
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

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
txtCodigo = ""
txtNome = ""
txtNascimento = "__/__/____"
txtRG = ""
txtCpf = "___.___.___-__"
cmbSexo.ListIndex = -1
txtCivil.ListIndex = -1
txtPIS = ""
txtReservista = ""
txtEleitor = ""
txtCTPS_N = ""
txtCTPS_serie = ""
txtDependente = ""
txtPai = ""
txtMae = ""
txtendereco = ""
txttelefone = ""
Cmb_centro.ListIndex = -1

Label57.Visible = True
fotopadrao = Localrel & "\Imagens\Caprind.bmp"
Foto.Picture = LoadPicture(fotopadrao)
CommonDialog1.filename = fotopadrao

1:
    txtFuncao = ""
    txttipo.ListIndex = -1
    txtData_adm = "__/__/____"
    txtSetor = ""
    txtDivisao = ""
    txtHorario_fim = "__:__"
    txtHorario_inicio = "__:__"
    txtTurno = ""
    txtIDfornecedor = ""
    txtFornecedor = ""
    txtConta_corrente = ""
    txtSalario = ""
    cmbSituacao.ListIndex = -1
    txtdata_desligado = "__/__/____"
    CodigoLista = 0
    Caption = "RH - Cadastro de funcionários"

Exit Sub
tratar_erro:
    If Err.Number = 76 Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Data = IIf(txtData = "", Date, txtData)
TBGravar!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
TBGravar!CODIGO = txtCodigo
TBGravar!Nome = txtNome
If IsDate(txtNascimento) = False Then
    TBGravar!data_nascimento = Null
    TBGravar!dia_nascimento = Null
Else
    TBGravar!data_nascimento = txtNascimento
    TBGravar!dia_nascimento = Day(txtNascimento)
End If
TBGravar!RG = txtRG
TBGravar!CPF = txtCpf
TBGravar!sexo = cmbSexo
TBGravar!Estado_civil = txtCivil
TBGravar!PIS = txtPIS
TBGravar!Reservista = txtReservista
TBGravar!Titulo_eleitor = txtEleitor
TBGravar!CTPS_N = txtCTPS_N
TBGravar!CTPS_serie = txtCTPS_serie
TBGravar!mae = txtMae
TBGravar!pai = txtPai
TBGravar!Endereco = txtendereco
TBGravar!telefone = txttelefone
If Cmb_centro <> "" Then TBGravar!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex) Else TBGravar!ID_CC = Null
TBGravar!Foto = CommonDialog1.filename

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 13, True
ProcCarregaToolBar2 Me, 15195, 8, True
ProcCarregaToolBar3 Me, 15195, 10, True
Formulario = "RH/Funcionários"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
USToolBar3.Visible = False
Aniversario = False
txtData_Inicio_Ferias.Value = Date
txtData_Fim_Ferias.Value = Date
txtInicio_Periodo_Ferias.Value = Date
txtFim_Periodo_Ferias.Value = Date
txtData_Sindicato.Value = Date
Cmb_emissao_ates.Value = Date
Cmb_validade_ates.Value = Date
txtData_Obs.Value = Date
Cmb_opcao_lista = "Validação"
ProcCarregaComboSetor Cmb_centro, "Setor is not null and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
ProcCarregaComboEmpresa Cmb_empresa, False

ProcRemoveObjetosResize Me
Top_Lista = Lista.Top
Height_Lista = Lista.Height

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "RH/Funcionários"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

If txtNascimento.Text <> "__/__/____" Then
    VerifData = txtNascimento.Text
    ProcVerificaData
    If VerifData = False Then
        txtNascimento.Text = "__/__/____"
        txtNascimento.SetFocus
        Exit Sub
    End If
End If
Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = True
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
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = True
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

Private Sub Imgcalendario2_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = True
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

Private Sub Imgcalendario3_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = True
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
Sit_Data = 4
FrmCalendario.Show 1

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
                If Cmb_opcao_lista = "Excluir" Then
                    If .ListItems(InitFor).SubItems(5) = "Sim" Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_ContasPagar", "int_codforn = " & .ListItems(InitFor) & " and Tipo = 'FU'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_contas_receber", "IDCliente = " & .ListItems(InitFor) & " and Tipo = 'FU'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "CFI", "Funcionario = '" & .ListItems(InitFor).ListSubItems(2) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                Else
                    If .ListItems(InitFor).SubItems(5) = "Não" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ID from funcionarios where ID = " & .ListItems(InitFor) & " and (Situacao IS NULL or Situacao = N'')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .ListItems.Item(InitFor).Checked = False
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                        TBAbrir.Close
                    End If
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

Private Sub Lista_doc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_doc
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_doc, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_doc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_doc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "este documento", "excluir", True, True) = False Then
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

Private Sub Lista_doc_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_doc.ListItems.Count = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")
TBMaterial.Open "Select * from Funcionarios_documentos where ID = " & Lista_doc.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    Proclimpacampos_doc
    txtID_doc = TBMaterial!ID
    txtData_doc = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
    txtResponsavel_doc = IIf(IsNull(TBMaterial!Responsavel), "", TBMaterial!Responsavel)
    Txt_caminho_doc = IIf(IsNull(TBMaterial!caminho), "", TBMaterial!caminho)
    Txt_obs_doc = IIf(IsNull(TBMaterial!Obs), "", TBMaterial!Obs)
    Novo_Funcionario8 = False
    Frame4.Enabled = True
    CodigoLista8 = Lista_doc.SelectedItem.index
End If
TBMaterial.Close

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
            If Cmb_opcao_lista = "Excluir" Then
                If .ListItems(InitFor).SubItems(5) = "Sim" Then
                    USMsgBox ("Não é possivel excluir funcionário, pois o mesmo esta validado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                Mensagem = "Não é permitido excluir este funcionário, pois o mesmo está sendo utilizado no módulo"
                ProcVerificaRegistroUtilizado "tbl_ContasPagar", "int_codforn = " & .ListItems(InitFor) & " and Tipo = 'FU' and Logsit = 'N'", "Financeiro/Contas a pagar"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_contas_receber", "IDCliente = " & .ListItems(InitFor) & " and Tipo = 'FU' and Logsit = 'N'", "Financeiro/Contas a receber"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_ContasPagar", "int_codforn = " & .ListItems(InitFor) & " and Tipo = 'FU' and Logsit = 'S'", "Financeiro/Contas pagas"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_contas_receber", "IDCliente = " & .ListItems(InitFor) & " and Tipo = 'FU' and Logsit = 'S'", "Financeiro/Contas recebidas"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "CFI", "Funcionario = '" & .ListItems(InitFor).ListSubItems(2) & "'", "Qualidade/Almoxarifado"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            Else
                If .ListItems(InitFor).SubItems(5) = "Não" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select ID from funcionarios where ID = " & .ListItems(InitFor) & " and (Situacao IS NULL or Situacao = N'')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido validar este funcionário, pois não foi informado a situação do mesmo."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBAbrir.Close
                        Exit Sub
                    End If
                    TBAbrir.Close
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from funcionarios where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
End If
TBAbrir.Close
CodigoLista = Lista.SelectedItem.index

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
txtId = TBAbrir!ID
txtData.Text = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
txtCodigo = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
txtNome = IIf(IsNull(TBAbrir!Nome), "", Trim(TBAbrir!Nome))
Caption = "RH - Cadastro de funcionários - Funcionário: " & txtNome
txtNascimento = IIf(IsNull(TBAbrir!data_nascimento), "__/__/____", Format(TBAbrir!data_nascimento, "dd/mm/yyyy"))
txtRG = IIf(IsNull(TBAbrir!RG), "", TBAbrir!RG)
txtCpf = IIf(IsNull(TBAbrir!CPF), "___.___.___-__", TBAbrir!CPF)
If IsNull(TBAbrir!sexo) = False And TBAbrir!sexo <> "" Then cmbSexo = TBAbrir!sexo
If IsNull(TBAbrir!Estado_civil) = False And TBAbrir!Estado_civil <> "" Then txtCivil = TBAbrir!Estado_civil
txtPIS = IIf(IsNull(TBAbrir!PIS), "", TBAbrir!PIS)
txtReservista = IIf(IsNull(TBAbrir!Reservista), "", TBAbrir!Reservista)
txtEleitor = IIf(IsNull(TBAbrir!Titulo_eleitor), "", TBAbrir!Titulo_eleitor)
txtCTPS_N = IIf(IsNull(TBAbrir!CTPS_N), "", TBAbrir!CTPS_N)
txtCTPS_serie = IIf(IsNull(TBAbrir!CTPS_serie), "", TBAbrir!CTPS_serie)
txtDependente = IIf(IsNull(TBAbrir!dependente), "", TBAbrir!dependente)
txtMae = IIf(IsNull(TBAbrir!mae), "", TBAbrir!mae)
txtPai = IIf(IsNull(TBAbrir!pai), "", TBAbrir!pai)
txtendereco = IIf(IsNull(TBAbrir!Endereco), "", TBAbrir!Endereco)
txttelefone = IIf(IsNull(TBAbrir!telefone), "", TBAbrir!telefone)

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Usuarios_setor.* from Usuarios_setor where ID = " & IIf(IsNull(TBAbrir!ID_CC), 0, TBAbrir!ID_CC), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
        Cmb_centro = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    Else
        Cmb_centro = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    End If
End If
TBFI.Close

1:
    If IsNull(TBAbrir!Foto) = False And TBAbrir!Foto <> "" Then
        Foto.Picture = LoadPicture(TBAbrir!Foto)
        CommonDialog1.filename = TBAbrir!Foto
        Label57.Visible = False
    Else
        Foto.Picture = LoadPicture(fotopadrao)
        CommonDialog1.filename = fotopadrao
        Label57.Visible = True
    End If
2:
    If IsNull(TBAbrir!Funcao) = False Then txtFuncao = TBAbrir!Funcao
    If IsNull(TBAbrir!Tipo) = False And TBAbrir!Tipo <> "" Then txttipo = TBAbrir!Tipo
    txtData_adm = IIf(IsNull(TBAbrir!Admissao), "__/__/____", Format(TBAbrir!Admissao, "dd/mm/yyyy"))
    txtSetor = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
    txtDivisao = IIf(IsNull(TBAbrir!divisao), "", TBAbrir!divisao)
    txtHorario_inicio = IIf(IsNull(TBAbrir!Horario_inicio), "__:__", Format(TBAbrir!Horario_inicio, "hh:mm"))
    txtHorario_fim = IIf(IsNull(TBAbrir!Horario_fim), "__:__", Format(TBAbrir!Horario_fim, "hh:mm"))
    txtTurno = IIf(IsNull(TBAbrir!Turno), "", TBAbrir!Turno)
    txtIDfornecedor = IIf(IsNull(TBAbrir!IDFornecedor), "", TBAbrir!IDFornecedor)
    txtConta_corrente = IIf(IsNull(TBAbrir!conta_corrente), "", TBAbrir!conta_corrente)
    txtSalario = IIf(IsNull(TBAbrir!Salario), "0,00", Format(TBAbrir!Salario, "###,##0.00"))
    If IsNull(TBAbrir!situacao) = False And TBAbrir!situacao <> "" Then cmbSituacao = TBAbrir!situacao
    txtdata_desligado = IIf(IsNull(TBAbrir!data_desligado), "__/__/____", Format(TBAbrir!data_desligado, "dd/mm/yyyy"))
    
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from compras_fornecedores where idcliente = " & IIf(txtIDfornecedor = "", 0, txtIDfornecedor), Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then txtFornecedor = IIf(IsNull(TBClientes!Nome_Razao), "", TBClientes!Nome_Razao)
    TBClientes.Close
    
    Frame2.Enabled = True
    Frame3.Enabled = True
    Novo_Funcionario = False
    ProcLimparTudo

Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        USMsgBox ("Não foi encontrado o centro de custo deste funcionário, favor alterar."), vbInformation, "CAPRIND v5.0"
        GoTo 1
    End If
    If Err.Number = "13" Or Err.Number = "53" Or Err.Number = "71" Or Err.Number = "75" Or Err.Number = "76" Then GoTo 2
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame11.Enabled = False
Frame12.Enabled = False
Frame13.Enabled = False
Frame14.Enabled = False
Frame15.Enabled = False
Frame1.Enabled = False
Frame16.Enabled = False
Frame4.Enabled = False
ProcLimpaCamposCurso
ProcLimpaCamposAumentos
ProcLimpaCamposFerias
ProcLimpaCamposSindicato
ProcLimpaCamposAtes
ProcLimpaCamposObs
Proclimpacampos_doc
ListaCursos.ListItems.Clear
ListaAumentos.ListItems.Clear
ListaFerias.ListItems.Clear
ListaSindicato.ListItems.Clear
ListaAtestados.ListItems.Clear
ListaObs.ListItems.Clear
Lista_doc.ListItems.Clear
Novo_funcionario1 = False
Novo_Funcionario2 = False
Novo_Funcionario3 = False
Novo_Funcionario4 = False
Novo_Funcionario5 = False
Novo_Funcionario6 = False
Novo_Funcionario7 = False
Novo_Funcionario8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

If StrSql_Localizar_Funcionarios = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_RH_Funcionarios = CreateObject("adodb.recordset")
TBLISTA_RH_Funcionarios.Open StrSql_Localizar_Funcionarios, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_RH_Funcionarios.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Cotacao = 0
Lista.ListItems.Clear
TBLISTA_RH_Funcionarios.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_RH_Funcionarios.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_RH_Funcionarios.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_RH_Funcionarios.RecordCount - IIf(Pagina > 1, (TBLISTA_RH_Funcionarios.PageSize * (Pagina - 1)), 0), TBLISTA_RH_Funcionarios.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_RH_Funcionarios.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_RH_Funcionarios!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_RH_Funcionarios!CODIGO), "", TBLISTA_RH_Funcionarios!CODIGO)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_RH_Funcionarios!Nome), "", Trim(TBLISTA_RH_Funcionarios!Nome))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_RH_Funcionarios!telefone), "", TBLISTA_RH_Funcionarios!telefone)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_RH_Funcionarios!Setor), "", TBLISTA_RH_Funcionarios!Setor)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_RH_Funcionarios!DtValidacao) = False, "Sim", "Não")
    End With
    TBLISTA_RH_Funcionarios.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_RH_Funcionarios.RecordCount
If TBLISTA_RH_Funcionarios.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_RH_Funcionarios.PageCount
ElseIf TBLISTA_RH_Funcionarios.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_RH_Funcionarios.PageCount & " de: " & TBLISTA_RH_Funcionarios.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_RH_Funcionarios.AbsolutePage - 1 & " de: " & TBLISTA_RH_Funcionarios.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaAumentos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaAumentos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaAumentos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaAumentos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaAumentos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "este aumento", "excluir", True, True) = False Then
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

Private Sub ListaCursos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaCursos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaCursos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCursos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaCursos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "este curso", "excluir", True, True) = False Then
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

Private Sub ListaCursos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaCursos.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposCurso
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from funcionarios_cursos where id = " & ListaCursos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIdCurso = TBAbrir!ID
    txtCurso = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    txtDtCurso = IIf(IsNull(TBAbrir!Data), "__/__/____", Format(TBAbrir!Data, "dd/mm/yyyy"))
    txtduracaocurso = IIf(IsNull(TBAbrir!duracao), "", TBAbrir!duracao)
    txtInstituicao = IIf(IsNull(TBAbrir!Instituicao), "", TBAbrir!Instituicao)
    txtNivel = IIf(IsNull(TBAbrir!Nivel), "", TBAbrir!Nivel)
    txt_Caminho = IIf(IsNull(TBAbrir!caminho), "", TBAbrir!caminho)
End If
1:
    TBAbrir.Close
    CodigoLista2 = ListaCursos.SelectedItem.index
    Frame11.Enabled = True
    Novo_Funcionario2 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaAumentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaAumentos.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposAumentos
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from funcionarios_aumentos where id = " & ListaAumentos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIdAumento = TBAbrir!ID
    txtMesAno = IIf(IsNull(TBAbrir!Data), "__/____", Format(TBAbrir!Data, "mm/yyyy"))
    txtValorAumento = IIf(IsNull(TBAbrir!valor), "", Format(TBAbrir!valor, "###,##0.00"))
    txtObsAumentos = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
End If
TBAbrir.Close
CodigoLista3 = ListaAumentos.SelectedItem.index
Frame12.Enabled = True
Novo_Funcionario3 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaFerias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaFerias
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaFerias, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaFerias_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaFerias
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "esta férias", "excluir", True, True) = False Then
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

Private Sub ListaFerias_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaFerias.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposFerias
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Funcionarios_ferias where id = " & ListaFerias.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIdFerias = TBAbrir!ID
    txtData_Inicio_Ferias.Value = TBAbrir!Data_inicio
    txtData_Fim_Ferias.Value = TBAbrir!Data_fim
    txtInicio_Periodo_Ferias.Value = TBAbrir!inicio_periodo
    txtFim_Periodo_Ferias.Value = TBAbrir!fim_periodo
End If
TBAbrir.Close
CodigoLista4 = ListaFerias.SelectedItem.index
Frame13.Enabled = True
Frame14.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaAtestados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaAtestados
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaAtestados, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaObs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaObs
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaObs, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaAtestados_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaAtestados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "este atestado", "excluir", True, True) = False Then
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

Private Sub ListaObs_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaObs
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "esta observação", "excluir", True, True) = False Then
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

Private Sub ListaSindicato_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaSindicato
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Funcionarios", "ID = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaSindicato, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaSindicato_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaSindicato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Funcionarios", "ID = " & txtId, "funcionario", "este sindicado", "excluir", True, True) = False Then
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

Private Sub ListaSindicato_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaSindicato.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposSindicato
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Funcionarios_sindicato where id = " & ListaSindicato.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIdSindicato = TBAbrir!ID
    txtData_Sindicato.Value = TBAbrir!Data
    txtValor_Sindicato = IIf(IsNull(TBAbrir!valor), "", Format(TBAbrir!valor, "###,##0.00"))
    Txt_obs_sindicato = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
End If
TBAbrir.Close
CodigoLista5 = ListaSindicato.SelectedItem.index
Frame15.Enabled = True
Novo_Funcionario5 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaAtestados_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaAtestados.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposAtes
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Funcionarios_atestados where ID = " & ListaAtestados.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Txt_ID_ates = TBAbrir!ID
    Txt_data_ates = Format(TBAbrir!Data, "dd/mm/yy")
    Txt_responsavel_ates = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    Cmb_emissao_ates.Value = TBAbrir!Data_emissao
    Cmb_tipo_ates = IIf(IsNull(TBAbrir!Tipo), "", IIf(TBAbrir!Tipo = "M", "Médico", "ASO"))
    Cmb_validade_ates.Value = TBAbrir!Data_validade
    Txt_caminho_ates = IIf(IsNull(TBAbrir!caminho), "", TBAbrir!caminho)
End If
TBAbrir.Close
CodigoLista6 = ListaAtestados.SelectedItem.index
Frame1.Enabled = True
Novo_Funcionario6 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaObs_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaObs.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposObs
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Funcionarios_obs where id = " & ListaObs.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIdObs = TBAbrir!ID
    txtData_Obs.Value = TBAbrir!Data
    txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
End If
TBAbrir.Close
CodigoLista7 = ListaObs.SelectedItem.index
Frame16.Enabled = True
Novo_Funcionario7 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then USToolBar3.Visible = False Else USToolBar3.Visible = True
ListaCursos.Visible = False
ListaAumentos.Visible = False
ListaFerias.Visible = False
ListaSindicato.Visible = False
ListaAtestados.Visible = False
ListaObs.Visible = False
Lista_doc.Visible = False
Frame5.Visible = False
Select Case SSTab1.Tab
    Case 0:
        Frame5.Visible = True
        With Lista
            .Visible = True
            .Top = Top_Lista
            .Height = Height_Lista
        End With
        Cmb_empresa.Visible = True
        Frame2.Visible = True
        
        If Lista.Visible = True Then Lista.SetFocus
    Case 1:
        Frame5.Visible = True
        Frame2.Visible = False

        With Lista
            .Visible = True
            .Top = Top_Lista - (Frame2.Height - Frame3.Height)
            .Height = Height_Lista + (Frame2.Height - Frame3.Height)
        End With
        Cmb_empresa.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista.SetFocus
    Case 2:
        ListaCursos.Visible = True
        Lista.Visible = False
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ListaCursos.SetFocus
        ProcAtualizalistaCursos
    Case 3:
        ListaAumentos.Visible = True
        Lista.Visible = False
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ListaAumentos.SetFocus
        ProcAtualizalistaAumentos
    Case 4:
        ListaFerias.Visible = True
        Lista.Visible = False
        ProcVerificaProsseguir
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        If Permitido = False Then Exit Sub
        ListaFerias.SetFocus
        ProcAtualizalistaFerias
    Case 5:
        ListaSindicato.Visible = True
        Lista.Visible = False
        ProcVerificaProsseguir
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        If Permitido = False Then Exit Sub
        ListaSindicato.SetFocus
        ProcAtualizalistaSindicato
    Case 6:
        ListaAtestados.Visible = True
        Lista.Visible = False
        ProcVerificaProsseguir
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        If Permitido = False Then Exit Sub
        ListaAtestados.SetFocus
        ProcAtualizalistaAtes
    Case 7:
        ListaObs.Visible = True
        Lista.Visible = False
        ProcVerificaProsseguir
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        If Permitido = False Then Exit Sub
        ListaObs.SetFocus
        ProcAtualizalistaObs
    Case 8:
        Lista_doc.Visible = True
        Lista.Visible = False
        ProcVerificaProsseguir
        Cmb_empresa.Visible = False
        Frame2.Visible = False
        If Permitido = False Then Exit Sub
        Lista_doc.SetFocus
        ProcAtualizaLista_Doc
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Funcionario = True Then
    USMsgBox ("Salve o funcionário antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    CmdSalvar.SetFocus
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtData_adm_LostFocus()
On Error GoTo tratar_erro

If txtData_adm.Text <> "__/__/____" Then
    VerifData = txtData_adm.Text
    ProcVerificaData
    If VerifData = False Then
        txtData_adm.Text = "__/__/____"
        txtData_adm.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdata_desligado_LostFocus()
On Error GoTo tratar_erro

If txtdata_desligado.Text <> "__/__/____" Then
    VerifData = txtdata_desligado.Text
    ProcVerificaData
    If VerifData = False Then
        txtdata_desligado.Text = "__/__/____"
        txtdata_desligado.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDependente_LostFocus()
On Error GoTo tratar_erro

If txtDependente.Text <> "" Then
    VerifNumero = txtDependente.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDependente.Text = ""
        txtDependente.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDtCurso_LostFocus()
On Error GoTo tratar_erro

If txtDtCurso.Text <> "__/__/____" Then
    VerifData = txtDtCurso.Text
    ProcVerificaData
    If VerifData = False Then
        txtDtCurso.Text = "__/__/____"
        txtDtCurso.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDfornecedor_LostFocus()
On Error GoTo tratar_erro

If txtIDfornecedor <> "" Then
    VerifNumero = txtIDfornecedor
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDfornecedor = ""
        txtIDfornecedor.SetFocus
        Exit Sub
    End If
    
    txtFornecedor = ""
    If txtIDfornecedor = "" Then Exit Sub
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from compras_fornecedores where IDCliente = " & txtIDfornecedor, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
        txtFornecedor = IIf(IsNull(TBClientes!Nome_Razao), "", TBClientes!Nome_Razao)
    End If
    TBClientes.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNascimento_LostFocus()
On Error GoTo tratar_erro

If txtNascimento.Text <> "__/__/____" Then
    VerifData = txtNascimento.Text
    ProcVerificaData
    If VerifData = False Then
        txtNascimento.Text = "__/__/____"
        txtNascimento.SetFocus
        Exit Sub
    End If
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

Private Sub txtSalario_Change()
On Error GoTo tratar_erro

If txtSalario <> "" Then
    VerifNumero = txtSalario
    ProcVerificaNumero
    If VerifNumero = False Then
        txtSalario = ""
        txtSalario.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSalario_LostFocus()
On Error GoTo tratar_erro

txtSalario = Format(txtSalario.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Sindicato_Change()
On Error GoTo tratar_erro

If txtValor_Sindicato <> "" Then
    VerifNumero = txtValor_Sindicato
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValor_Sindicato = ""
        txtValor_Sindicato.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Sindicato_LostFocus()
On Error GoTo tratar_erro

txtValor_Sindicato = Format(txtValor_Sindicato, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorAumento_Change()
On Error GoTo tratar_erro

If txtValorAumento <> "" Then
    VerifNumero = txtValorAumento
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValorAumento = ""
        txtValorAumento.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorAumento_LostFocus()
On Error GoTo tratar_erro

txtValorAumento = Format(txtValorAumento, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Funcionarios order by Nome", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimpaCamposAumentos
        ProcLimpaCamposCurso
        ProcLimpaCamposFerias
        ProcLimpaCamposObs
        ProcLimpaCamposSindicato
        Proclimpacampos_doc
        txtId.Text = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Funcionarios where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        TBAbrir.Close
        ProcAtualizalistaAumentos
        ProcAtualizalistaCursos
        ProcAtualizalistaFerias
        ProcAtualizalistaObs
        ProcAtualizalistaSindicato
        ProcAtualizaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros de funcionários."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Funcionario = False
Novo_Funcionario2 = False
Novo_Funcionario3 = False
Novo_Funcionario4 = False
Novo_Funcionario5 = False
Novo_Funcionario7 = False
Novo_Funcionario8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Funcionarios order by Nome", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCamposAumentos
        ProcLimpaCamposCurso
        ProcLimpaCamposFerias
        ProcLimpaCamposObs
        ProcLimpaCamposSindicato
        Proclimpacampos_doc
        txtId.Text = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Funcionarios where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        TBAbrir.Close
        ProcAtualizalistaAumentos
        ProcAtualizalistaCursos
        ProcAtualizalistaFerias
        ProcAtualizalistaObs
        ProcAtualizalistaSindicato
        ProcAtualizaLista_Doc
    Else
        USMsgBox ("Fim dos cadastros de funcionários."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Funcionario = False
Novo_Funcionario2 = False
Novo_Funcionario3 = False
Novo_Funcionario4 = False
Novo_Funcionario5 = False
Novo_Funcionario7 = False
Novo_Funcionario8 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcCarregarTodos
    Case 9: ProcValidarRegistros Lista, "RH/Funcionários"
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar1
    Case 2: ProcImprimir
    Case 3: ProcAnterior
    Case 4: ProcProximo
    Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 2:
        Select Case ButtonIndex
            Case 1: ProcNovoCurso
            Case 2: ProcSalvarCurso
            Case 3: ProcExcluirCurso
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 3:
        Select Case ButtonIndex
            Case 1: ProcNovoAumento
            Case 2: ProcSalvarAumento
            Case 3: ProcExcluirAumento
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 4:
        Select Case ButtonIndex
            Case 1: ProcNovoFerias
            Case 2: ProcSalvarFerias
            Case 3: ProcExcluirFerias
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 5:
        Select Case ButtonIndex
            Case 1: ProcNovoSindicato
            Case 2: ProcSalvarSindicato
            Case 3: ProcExcluirSindicato
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 6:
        Select Case ButtonIndex
            Case 1: ProcNovoAtes
            Case 2: ProcSalvarAtes
            Case 3: ProcExcluirAtes
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 7:
        Select Case ButtonIndex
            Case 1: ProcNovoObs
            Case 2: ProcSalvarObs
            Case 3: ProcExcluirObs
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
    Case 8:
        Select Case ButtonIndex
            Case 1: procNovo_doc
            Case 2: procSalvar_doc
            Case 3: procExcluir_doc
            Case 4: ProcImprimir
            Case 5: ProcAnterior
            Case 6: ProcProximo
            Case 8: ProcAjuda
            Case 9: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
