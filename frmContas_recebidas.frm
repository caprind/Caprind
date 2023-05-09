VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContas_recebidas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Contas recebidas"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmContas_recebidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
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
   Begin VB.TextBox txt_tituloref 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   3150
      MaxLength       =   20
      TabIndex        =   103
      Top             =   5700
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtidintconta 
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
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      ToolTipText     =   "Número da conta."
      Top             =   5700
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   94
      Top             =   8775
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
         ItemData        =   "frmContas_recebidas.frx":212A
         Left            =   6960
         List            =   "frmContas_recebidas.frx":2134
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   32
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   36
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_recebidas.frx":214C
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
         TabIndex        =   35
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_recebidas.frx":58F0
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
         TabIndex        =   33
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
         TabIndex        =   34
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_recebidas.frx":93F9
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
         TabIndex        =   37
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_recebidas.frx":D4E8
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
         Index           =   1
         Left            =   3360
         TabIndex        =   107
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label21 
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
         Index           =   1
         Left            =   5610
         TabIndex        =   105
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label16 
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
         Left            =   2040
         TabIndex        =   104
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
         TabIndex        =   96
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
         TabIndex        =   95
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   75
      TabIndex        =   78
      Top             =   9390
      Width           =   15195
      Begin VB.TextBox TXTTOTAL 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Total geral recebido."
         Top             =   180
         Width           =   1620
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   91
         Top             =   210
         Width           =   11955
         _ExtentX        =   21087
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral :"
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
         Left            =   12270
         TabIndex        =   61
         Top             =   180
         Width           =   2445
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   75
      TabIndex        =   93
      Top             =   330
      Width           =   15210
      _ExtentX        =   26829
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
      ButtonLeft2     =   46
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   92
      ButtonTop3      =   2
      ButtonWidth3    =   45
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
      ButtonLeft4     =   139
      ButtonTop4      =   2
      ButtonWidth4    =   60
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "C. contábil"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Conta contábil (F6)"
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
      ButtonLeft5     =   201
      ButtonTop5      =   2
      ButtonWidth5    =   66
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Cancelar recompra"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Cancelar recompra (F7)"
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
      ButtonLeft6     =   269
      ButtonTop6      =   2
      ButtonWidth6    =   115
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Visualizar"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Visualizar antecipações relacionadas (F8)"
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
      ButtonLeft7     =   386
      ButtonTop7      =   2
      ButtonWidth7    =   62
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonAlignment8=   2
      ButtonType8     =   1
      ButtonStyle8    =   -1
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   -1
      ButtonLeft8     =   450
      ButtonTop8      =   4
      ButtonWidth8    =   2
      ButtonHeight8   =   54
      ButtonCaption9  =   "Ajuda"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Ajuda (F1)"
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
      ButtonLeft9     =   454
      ButtonTop9      =   2
      ButtonWidth9    =   41
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Sair"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Sair (Esc)"
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
      ButtonLeft10    =   497
      ButtonTop10     =   2
      ButtonWidth10   =   30
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonKey11     =   "11"
      ButtonAlignment11=   2
      BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState11   =   5
      ButtonLeft11    =   529
      ButtonTop11     =   2
      ButtonWidth11   =   24
      ButtonHeight11  =   24
      ButtonUseMaskColor11=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   8820
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmContas_recebidas.frx":10D74
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3765
      Left            =   75
      TabIndex        =   29
      Top             =   4995
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6641
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. venc."
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   4860
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. baixado"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Dt. baixa"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "N. boleto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Remessa"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar contas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   75
      TabIndex        =   88
      Top             =   4455
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
         ItemData        =   "frmContas_recebidas.frx":1714B
         Left            =   14220
         List            =   "frmContas_recebidas.frx":1714D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         ToolTipText     =   "Ano."
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton OptDomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   57
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton OptAteomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   58
         Top             =   270
         Width           =   1035
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   375
         Left            =   2130
         TabIndex        =   59
         Top             =   240
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   661
         MultiRow        =   -1  'True
         TabMinWidth     =   1439
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   13
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Janeiro"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fevereiro"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Março"
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
               Caption         =   "Junho"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Julho"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Agosto"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Setembro"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Outubro"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Novembro"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dezembro"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Todas"
               Key             =   "Todas"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Dados da conta"
      TabPicture(0)   =   "frmContas_recebidas.frx":1714F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados da baixa"
      TabPicture(1)   =   "frmContas_recebidas.frx":1716B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   3135
         Left            =   75
         TabIndex        =   62
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_localizar_contatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   10260
            Picture         =   "frmContas_recebidas.frx":17187
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Localizar contatos."
            Top             =   990
            Width           =   315
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
            Picture         =   "frmContas_recebidas.frx":1749B
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Filtrar por valor."
            Top             =   990
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
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmContas_recebidas.frx":178B6
            Left            =   5175
            List            =   "frmContas_recebidas.frx":178B8
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Tipo do documento."
            Top             =   370
            Width           =   1065
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
            Left            =   6225
            Picture         =   "frmContas_recebidas.frx":178BA
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Filtrar por tipo do documento."
            Top             =   370
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
            Left            =   4770
            Picture         =   "frmContas_recebidas.frx":17CD5
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Filtrar por data da transação."
            Top             =   370
            Width           =   315
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
            Height          =   855
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            ToolTipText     =   "Observações."
            Top             =   2130
            Width           =   7230
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
            ItemData        =   "frmContas_recebidas.frx":180F0
            Left            =   1755
            List            =   "frmContas_recebidas.frx":18100
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Tipo."
            Top             =   990
            Width           =   1905
         End
         Begin VB.TextBox txtParcela 
            BackColor       =   &H00FFFFFF&
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
            Left            =   14295
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Número da parcela."
            Top             =   370
            Width           =   720
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
            ItemData        =   "frmContas_recebidas.frx":1813C
            Left            =   180
            List            =   "frmContas_recebidas.frx":1813E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   370
            Width           =   3375
         End
         Begin VB.ComboBox txtstatus 
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
            Height          =   330
            ItemData        =   "frmContas_recebidas.frx":18140
            Left            =   180
            List            =   "frmContas_recebidas.frx":1815C
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Status."
            Top             =   1545
            Width           =   6900
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
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Número do documento."
            Top             =   370
            Width           =   1500
         End
         Begin VB.CommandButton Cmdlocalizarcliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9930
            Picture         =   "frmContas_recebidas.frx":18248
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar cliente."
            Top             =   990
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
            Picture         =   "frmContas_recebidas.frx":1834A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Filtrar por nome do cliente."
            Top             =   990
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
            Left            =   9225
            Picture         =   "frmContas_recebidas.frx":18765
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Filtrar por número da nota fiscal."
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
            Left            =   10710
            Picture         =   "frmContas_recebidas.frx":18B80
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Filtrar por número do pedido interno."
            Top             =   370
            Width           =   315
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
            Left            =   12300
            Picture         =   "frmContas_recebidas.frx":18F9B
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Filtrar por data de emissão."
            Top             =   370
            Width           =   315
         End
         Begin VB.TextBox txtuf 
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
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   420
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
            Left            =   7095
            Picture         =   "frmContas_recebidas.frx":193B6
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Filtrar por status."
            Top             =   1545
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
            Left            =   13905
            Picture         =   "frmContas_recebidas.frx":197D1
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Filtrar por data de vencimento."
            Top             =   370
            Width           =   315
         End
         Begin VB.TextBox txtNFiscal 
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
            Left            =   8130
            MaxLength       =   50
            TabIndex        =   6
            ToolTipText     =   "Número da nota fiscal."
            Top             =   370
            Width           =   1095
         End
         Begin VB.ComboBox txtProposta 
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
            Left            =   9615
            TabIndex        =   8
            ToolTipText     =   "Número do pedido interno."
            Top             =   370
            Width           =   1095
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
            TabIndex        =   18
            ToolTipText     =   "Código do cliente."
            Top             =   990
            Width           =   810
         End
         Begin VB.TextBox txtNome_Razao 
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
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Nome do cliente."
            Top             =   990
            Width           =   5115
         End
         Begin VB.TextBox txtCidade 
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
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Cidade."
            Top             =   990
            Width           =   3900
         End
         Begin MSComCtl2.DTPicker mskEmissao 
            Height          =   315
            Left            =   11100
            TabIndex        =   10
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
            Format          =   182059011
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker mskVencimento 
            Height          =   315
            Left            =   12690
            TabIndex        =   12
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
            Format          =   182059011
            CurrentDate     =   39057
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Valor."
            Top             =   990
            Width           =   1170
         End
         Begin MSComctlLib.ListView Lista_PC 
            Height          =   1445
            Left            =   7515
            TabIndex        =   28
            Top             =   1545
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   2540
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
            Left            =   3570
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
            Format          =   182059011
            CurrentDate     =   39057
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo docto."
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
            Left            =   5295
            TabIndex        =   102
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. transação"
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
            Left            =   3675
            TabIndex        =   101
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   3135
            TabIndex        =   99
            Top             =   1920
            Width           =   1050
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
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
            Index           =   0
            Left            =   10530
            TabIndex        =   98
            Top             =   1350
            Width           =   1470
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   2550
            TabIndex        =   97
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parcela"
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
            Index           =   2
            Left            =   14363
            TabIndex        =   92
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Left            =   1500
            TabIndex        =   90
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº documento"
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
            Height          =   180
            Left            =   6855
            TabIndex        =   81
            Top             =   180
            Width           =   1020
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
            Left            =   3353
            TabIndex        =   73
            Top             =   1350
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nota fiscal"
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
            Height          =   180
            Left            =   8295
            TabIndex        =   70
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ped. interno"
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
            Height          =   180
            Index           =   0
            Left            =   9720
            TabIndex        =   69
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. emissão"
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
            Height          =   180
            Index           =   1
            Left            =   11280
            TabIndex        =   68
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. vencto."
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
            Height          =   180
            Left            =   12870
            TabIndex        =   67
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Left            =   540
            TabIndex        =   66
            Top             =   780
            Width           =   450
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
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
            Left            =   6810
            TabIndex        =   65
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
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
            Left            =   12383
            TabIndex        =   64
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
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
            Left            =   14708
            TabIndex        =   63
            Top             =   780
            Width           =   195
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   2055
         Left            =   -74925
         TabIndex        =   71
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_valor_recebido 
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
            Left            =   1395
            Picture         =   "frmContas_recebidas.frx":19BEC
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Filtrar por valor pago."
            Top             =   370
            Width           =   315
         End
         Begin VB.TextBox Txt_local_desconto 
            BackColor       =   &H00FFFFFF&
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
            Left            =   7315
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Local do desconto."
            Top             =   1590
            Width           =   7695
         End
         Begin VB.ComboBox txtFormaPagto 
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
            ItemData        =   "frmContas_recebidas.frx":1A007
            Left            =   10995
            List            =   "frmContas_recebidas.frx":1A009
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   49
            ToolTipText     =   "Forma da baixa."
            Top             =   370
            Width           =   4020
         End
         Begin VB.TextBox txt_Ndocto 
            BackColor       =   &H00FFFFFF&
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
            Left            =   13505
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Número do documento baixa."
            Top             =   990
            Width           =   1505
         End
         Begin VB.TextBox txtdesconto 
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
            Left            =   9735
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Valor do desconto."
            Top             =   370
            Width           =   1245
         End
         Begin VB.TextBox txtjuros 
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
            Left            =   5955
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Valor diário do juros de mora."
            Top             =   370
            Width           =   1245
         End
         Begin VB.TextBox Txt_dias_atraso 
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
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Dias em atraso."
            Top             =   370
            Width           =   1125
         End
         Begin VB.TextBox Txt_multa 
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
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Valor da multa."
            Top             =   370
            Width           =   1245
         End
         Begin VB.TextBox Txt_total_juros 
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
            Left            =   7215
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de juros."
            Top             =   370
            Width           =   1245
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
            Left            =   11940
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "N° da conta corrente."
            Top             =   990
            Width           =   1550
         End
         Begin VB.TextBox txtobservacao_recbto 
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
            Height          =   915
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            ToolTipText     =   "Observações do recebimento."
            Top             =   990
            Width           =   7095
         End
         Begin VB.CommandButton cmdrecebimento 
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
            Left            =   3090
            Picture         =   "frmContas_recebidas.frx":1A00B
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Filtrar por data de recebimento."
            Top             =   370
            Width           =   315
         End
         Begin VB.TextBox txtvalortitrecebido 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Valor baixado."
            Top             =   370
            Width           =   1200
         End
         Begin VB.CheckBox chbparcial 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bx. parcial"
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
            Height          =   195
            Left            =   270
            TabIndex        =   50
            Top             =   780
            Width           =   1185
         End
         Begin VB.CommandButton cmdinstituicao 
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
            Left            =   11550
            Picture         =   "frmContas_recebidas.frx":1A426
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Filtrar por instituição bancária."
            Top             =   990
            Width           =   315
         End
         Begin VB.ComboBox txtBanco 
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
            Left            =   7320
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   52
            ToolTipText     =   "Instituição bancária."
            Top             =   990
            Width           =   4215
         End
         Begin MSComCtl2.DTPicker mskData_pagamento 
            Height          =   315
            Left            =   1815
            TabIndex        =   41
            ToolTipText     =   "Data de recebimento."
            Top             =   375
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   182190083
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker Cmb_data_movimentacao 
            Height          =   315
            Left            =   3510
            TabIndex        =   43
            ToolTipText     =   "Data da movimentação."
            Top             =   375
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   182190081
            CurrentDate     =   39057
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. moviment."
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
            Left            =   3540
            TabIndex        =   106
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local do desconto"
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
            Left            =   10522
            TabIndex        =   89
            Top             =   1380
            Width           =   1290
         End
         Begin VB.Label LblDocumento 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "N° documento baixa"
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
            Left            =   13505
            TabIndex        =   87
            Top             =   780
            Width           =   1470
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto"
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
            Index           =   2
            Left            =   10020
            TabIndex        =   86
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Juros diário"
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
            Left            =   6165
            TabIndex        =   85
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias em atraso"
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
            Left            =   4845
            TabIndex        =   84
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Multa"
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
            Left            =   8895
            TabIndex        =   83
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total de juros"
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
            TabIndex        =   82
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma da baixa"
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
            Index           =   3
            Left            =   12450
            TabIndex        =   80
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta corrente"
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
            Left            =   12173
            TabIndex        =   79
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Observações do recebimento"
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
            Left            =   2655
            TabIndex        =   77
            Top             =   780
            Width           =   2145
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. baixa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2130
            TabIndex        =   76
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. baixado"
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
            Left            =   300
            TabIndex        =   75
            Top             =   180
            Width           =   990
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
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
            Left            =   9210
            TabIndex        =   74
            Top             =   780
            Width           =   435
         End
      End
   End
End
Attribute VB_Name = "frmContas_recebidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrSql_Contas_Recebidas As String 'OK
Public StrSql_Contas_RecebidasTotal As String 'OK
Public FormulaRel_Contas_Recebidas As String 'OK
Dim TBLISTA_Contas_Recebidas As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=3l3vDpTucL0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=12&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarTodas()
On Error GoTo tratar_erro

ProcConstruirFiltroPadrao "CR.IDintconta IS NOT NULL", "Not(IsNull({tbl_Contas_receber.IDintconta}))", "Data_pagamento"
ProcSalvarDadosRel False, False, False, Date, Date, Date
ProcCarregaLista (1)
Imprimir = True

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
Set TBLISTA_Contas_Recebidas = CreateObject("adodb.recordset")
TBLISTA_Contas_Recebidas.Open StrSql_Contas_Recebidas, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Contas_Recebidas.EOF = False Then ProcExibePagina (Pagina)
ProcCarregaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Codproduto = 0
ValorTotal = 0
Lista.ListItems.Clear
TBLISTA_Contas_Recebidas.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Contas_Recebidas.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Contas_Recebidas.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Contas_Recebidas.RecordCount - IIf(Pagina > 1, (TBLISTA_Contas_Recebidas.PageSize * (Pagina - 1)), 0), TBLISTA_Contas_Recebidas.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Contas_Recebidas.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add.Text = TBLISTA_Contas_Recebidas!IDintconta
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Contas_Recebidas!emissao), "", Format(TBLISTA_Contas_Recebidas!emissao, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Contas_Recebidas!Vencimento), "", Format(TBLISTA_Contas_Recebidas!Vencimento, "dd/mm/yy"))
        
        If TBLISTA_Contas_Recebidas!Parcial = True And TBLISTA_Contas_Recebidas!status <> "TÍTULO RECEBIDO PARCIAL LIQUIDADO" Then
            valor = IIf(IsNull(TBLISTA_Contas_Recebidas!RecebidoParcial), 0, TBLISTA_Contas_Recebidas!RecebidoParcial) + IIf(IsNull(TBLISTA_Contas_Recebidas!ValorPendente), 0, TBLISTA_Contas_Recebidas!ValorPendente)
        Else
            valor = IIf(IsNull(TBLISTA_Contas_Recebidas!valor), 0, TBLISTA_Contas_Recebidas!valor)
        End If
        .Item(.Count).SubItems(3) = Format(valor, "###,##0.00")
        
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Contas_Recebidas!txt_ndocumento), "", TBLISTA_Contas_Recebidas!txt_ndocumento)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Contas_Recebidas!NFiscal), "", TBLISTA_Contas_Recebidas!NFiscal)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Contas_Recebidas!Parcela), "", TBLISTA_Contas_Recebidas!Parcela)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Contas_Recebidas!Nome_Razao), "", Trim(TBLISTA_Contas_Recebidas!Nome_Razao))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Contas_Recebidas!valortitulorecebido), "", Format(TBLISTA_Contas_Recebidas!valortitulorecebido, "###,##0.00"))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Contas_Recebidas!Data_pagamento), "", Format(TBLISTA_Contas_Recebidas!Data_pagamento, "dd/mm/yy"))
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Contas_Recebidas!resprec), "", TBLISTA_Contas_Recebidas!resprec)
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select I.Id as ID_banco, DR.* from tbl_Detalhes_Recebimento DR INNER JOIN tbl_Instituicoes I ON DR.txt_Portador_Banco = I.txt_Descricao where DR.IDContaReceber = " & TBLISTA_Contas_Recebidas!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(11) = IIf(IsNull(TBAbrir!Nosso_Numero), "", TBAbrir!Nosso_Numero)
            If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" And TBAbrir!ID_banco <> "" Then .Item(.Count).SubItems(12) = FunFormataNumeroArqRemessa(TBAbrir!Data_emissao, TBAbrir!ID_banco, TBAbrir!Seq_remessa)
        Else
            .Item(.Count).SubItems(11) = ""
        End If
        TBAbrir.Close
    End With
    TBLISTA_Contas_Recebidas.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Contas_Recebidas.RecordCount
If TBLISTA_Contas_Recebidas.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Contas_Recebidas.PageCount
ElseIf TBLISTA_Contas_Recebidas.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Recebidas.PageCount & " de: " & TBLISTA_Contas_Recebidas.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Recebidas.AbsolutePage - 1 & " de: " & TBLISTA_Contas_Recebidas.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

ValorTotal = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open StrSql_Contas_RecebidasTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    ValorTotal = IIf(IsNull(TBTotaisnota!TotContas), 0, TBTotaisnota!TotContas)
End If
TBTotaisnota.Close
txtTotal.Text = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

ProcCorrigeLayoutForm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcCarregaComboBanco

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelarRecompra()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtidintconta = "" Then
    NomeCampo = "a conta"
    Acao = "cancelar a recompra"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus <> "DUPLICATA DESCONTADA RECOMPRADA" Then
    USMsgBox ("Esta duplicata não está recomprada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente cancelar a recompra desta duplicata?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_contas_receber where IdContaRecomprada = " & txtidintconta & " and status <> 'TÍTULO EM ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        USMsgBox ("Não é permitido cancelar esta recompra, pois esta duplicata está vinculada a uma conta a receber com o status diferente de título em aberto."), vbExclamation, "CAPRIND v5.0"
        TBContas.Close
        Exit Sub
    End If
    TBContas.Close
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select IDduplicata, valor_enviado from troca_titulo_valores where n_conta = " & txtidintconta.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select local_troca from troca_titulo where id = " & TBAbrir!IDduplicata, Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Set TBReceber = CreateObject("adodb.recordset")
            TBReceber.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & TBContas!local_troca & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBReceber.EOF = False Then
                'Verifica se alterou o status da conta a pagar
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from tbl_ContasPagar where IdContaReceber = " & txtidintconta & " and Status <> 'TÍTULO EM ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    USMsgBox ("Não é permitido alterar o status desta conta, pois o status da conta à pagar que está amarrada a esta foi alterado."), vbExclamation, "CAPRIND v5.0"
                    TBFI.Close
                    Exit Sub
                End If
                TBFI.Close
                'Verifica limite de desconto no banco
                valor = Format(TBAbrir!valor_enviado, "###,##0.00")
                If valor + TBReceber!Limite_utilizado > TBReceber!Limite_desconto Then
                    USMsgBox ("Não é permitido alterar o status deste título pois o limite para desconto hoje é de " & Format(TBReceber!Limite_desconto - TBReceber!Limite_utilizado, "###,##0.00") & "."), vbExclamation, "CAPRIND v5.0"
                    If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                        GoTo Prosseguir
                    Else
                        TBReceber.Close
                        Exit Sub
                    End If
                End If
Prosseguir:
                TBReceber!Limite_utilizado = Format(TBReceber!Limite_utilizado + valor, "###,##0.00")
                TBReceber.Update
            End If
            TBReceber.Close
            
            ProcExcluiContaPagar
            
            'Exclui a conta a receber cadastrada atravez da recompra
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_contas_receber where IdContaRecomprada = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Conexao.Execute "DELETE familia_financeiro where IDConta = " & TBFI!IDintconta & " and TipoConta = 'R'"
                
                'Fluxo de Caixa
                Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBFI!IDFluxo), 0, TBFI!IDFluxo)
                
                TBFI.Delete
            End If
            TBFI.Close
            
            Conexao.Execute "Update tbl_contas_receber Set LogSit = 'N', Bloqueado = 'False', Status = 'DUPLICATA DESCONTADA EM ABERTO', data_pagamento = Null, resprec = Null where IdIntConta = " & txtidintconta
            Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where IDConta = " & txtidintconta & " and TipoConta = 'R'"
            
            
       End If
       TBContas.Close
    End If
    TBAbrir.Close
    USMsgBox ("Recompra cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Financeiro/Contas recebidas"
    Evento = "Cancelar recompra"
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

Private Sub ProcExcluiContaPagar()
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_ContasPagar where IDContaReceber = " & txtidintconta.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Conexao.Execute "DELETE familia_financeiro where IDConta = " & TBFI!IDintconta & " and TipoConta = 'P'"
    
    'Fluxo de Caixa
    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBFI!IDFluxo), 0, TBFI!IDFluxo)
    
    TBFI.Delete
End If
'Grava valor de recompra no bordero
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select IDduplicata, valor_enviado from troca_titulo_valores where n_conta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select Vlrtotalrecompra from troca_titulo where id = " & TBFI!IDduplicata, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar!Vlrtotalrecompra = IIf(IsNull(TBGravar!Vlrtotalrecompra), 0, TBGravar!Vlrtotalrecompra) - TBFI!valor_enviado
        TBGravar.Update
    End If
    TBGravar.Close
End If
TBFI.Close

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
    If Cmb_opcao_lista = "Excluir" Then .ButtonState(3) = 0 Else .ButtonState(3) = 5
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

ProcFiltrarContas "Data_transacao = '" & Format(Txt_data_transacao.Value, "Short Date") & "'", "{tbl_Contas_receber.Data_transacao} = Date(" & Year(Txt_data_transacao.Value) & "," & Month(Txt_data_transacao.Value) & "," & Day(Txt_data_transacao.Value) & ")", True, True, False, False, False, Txt_data_transacao, Txt_data_transacao, "Data_transacao"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcFiltrarContas(TextoFiltro As String, TextoFiltroRel As String, Imprimir As Boolean, DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataBaixa As Boolean, DataInicio As Date, DataFinal As Date, Ordenar As String)
On Error GoTo tratar_erro

ProcConstruirFiltroPadrao TextoFiltro, TextoFiltroRel, Ordenar
ProcSalvarDadosRel DataTransacao, DataEmissao, DataVencimento, DataBaixa, DataInicio, DataFinal
ProcCarregaLista (1)
Imprimir = Imprimir

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcConstruirFiltroPadrao(TextoFiltro As String, TextoFiltroRel As String, Ordenar As String)
On Error GoTo tratar_erro

CamposFiltro = "CR.IDintconta, CR.emissao, CR.Vencimento, CR.Data_transacao, CR.Parcial, CR.Status, CR.RecebidoParcial, CR.ValorPendente, CR.Valor, CR.txt_ndocumento, CR.NFiscal, CR.Parcela, CR.Nome_Razao, CR.valortitulorecebido, CR.Data_pagamento, CR.resprec"
If Left(TextoFiltro, 2) = "PN" Then INNERJOINPADRAO = " from tbl_contas_receber CR INNER JOIN tbl_proposta_nota PN ON PN.ID_nota = CR.ID_nota" Else INNERJOINPADRAO = " from tbl_contas_receber CR"
INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
INNERJOINTEXTOSUM = "Select Sum(CR.valortitulorecebido) as TotContas " & INNERJOINPADRAO
TextoFiltroPadrao = "CR.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and CR.LogSit = 'S'"
TextoFiltroPadraoRel = "{tbl_contas_receber.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_contas_receber.LogSit} = 'S'"
OrdenarTexto = " group by " & CamposFiltro & " order by CR." & Ordenar & " desc, CR.IdIntConta"
StrSql_Contas_Recebidas = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
StrSql_Contas_RecebidasTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadrao
FormulaRel_Contas_Recebidas = TextoFiltroRel & " and " & TextoFiltroPadraoRel

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
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = True
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

Private Sub Cmd_valor_Click()
On Error GoTo tratar_erro

If txtValor <> "" Then
    valor = txtValor
    NovoValor = Replace(valor, ",", ".")
    ProcFiltrarContas "Valor = " & NovoValor, "{tbl_Contas_receber.Valor} = " & NovoValor, True, False, False, False, False, Date, Date, "data_pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_valor_recebido_Click()
On Error GoTo tratar_erro

If txtvalortitrecebido <> "" Then
    valor = txtvalortitrecebido
    NovoValor = Replace(valor, ",", ".")
    ProcFiltrarContas "Valortitulorecebido = " & NovoValor, "{tbl_Contas_receber.Valortitulorecebido} = " & NovoValor, True, False, False, False, False, Date, Date, "data_pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmddoc_Click()
On Error GoTo tratar_erro

If txtNFiscal.Text <> "" Then
    ProcFiltrarContas "nfiscal = '" & txtNFiscal.Text & "'", "{tbl_Contas_receber.nfiscal} = '" & txtNFiscal.Text & "'", True, False, False, False, False, Date, Date, "data_pagamento"
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

ProcFiltrarContas "emissao = '" & Format(mskEmissao.Value, "Short Date") & "'", "{tbl_Contas_receber.emissao} = Date(" & Year(mskEmissao.Value) & "," & Month(mskEmissao.Value) & "," & Day(mskEmissao.Value) & ")", True, False, True, False, False, mskEmissao, mskEmissao, "Emissao"

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

txtBanco.ListIndex = -1
txtFormaPagto.ListIndex = -1
cmbtipo_conta.ListIndex = -1
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_contas_receber where IdIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!Banco) = False And TBFI!Banco <> "" Then txtBanco = TBFI!Banco
    If IsNull(TBFI!FormaBaixa) = False And TBFI!FormaBaixa <> "" Then txtFormaPagto = TBFI!FormaBaixa
    If IsNull(TBFI!Tipo_doc) = False And TBFI!Tipo_doc <> "" Then cmbtipo_conta = TBFI!Tipo_doc
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
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
mskVencimento.Value = Date
txtValor.Text = ""
txtparcela.Text = ""
txtIDcliente.Text = ""
txtNome_Razao.Text = ""
txtuf.Text = ""
txtCidade.Text = ""
txtStatus.ListIndex = -1
txtobservacao_recbto.Text = ""
Txt_local_desconto = ""
txtFormaPagto.ListIndex = -1
mskData_pagamento.Value = Date
Cmb_data_movimentacao.Value = Date
chbparcial.Value = False
txtvalortitrecebido.Text = ""
txt_Ndocto = ""
Txt_dias_atraso = ""
txtjuros = ""
Txt_total_juros = ""
Txt_multa = ""
txtDesconto = ""
txtBanco.ListIndex = -1
txtConta = ""
txt_tituloref.Text = ""
CodigoLista = 0
Lista_PC.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPlanoContas()
On Error GoTo tratar_erro
    
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Recebidas = True
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

Private Sub ProcAntecipacoes()
On Error GoTo tratar_erro

If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de visualizar a lista de antecipações/devoluções relacionadas."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Financeiro_Contas_Pagas = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_contas_antecipacao where ID_conta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Or txtStatus = "TÍTULO LIQUIDADO ANTECIPADO" Then
    frmContas_antecipacoes.Show 1
Else
    frmContas_devolucoes.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdinstituicao_Click()
On Error GoTo tratar_erro

If txtBanco.Text <> "" Then
    ProcFiltrarContas "banco = '" & txtBanco.Text & "'", "{tbl_Contas_receber.banco} = '" & txtBanco.Text & "'", True, False, False, False, False, Date, Date, "data_pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_fornecedor_Click()
On Error GoTo tratar_erro

If txtNome_Razao.Text <> "" Then
    ProcFiltrarContas "nome_razao = '" & txtNome_Razao.Text & "'", "{tbl_Contas_receber.nome_razao} = '" & txtNome_Razao.Text & "'", True, False, False, False, False, Date, Date, "data_pagamento"
Else
    ProcFiltrarTodas
End If

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

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False
ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False
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

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Recebidas.AbsolutePage <> 2 Then
   If TBLISTA_Contas_Recebidas.AbsolutePage = -3 Then
      ProcExibePagina (TBLISTA_Contas_Recebidas.PageCount - 1)
   Else
      TBLISTA_Contas_Recebidas.AbsolutePage = TBLISTA_Contas_Recebidas.AbsolutePage - 2
      ProcExibePagina (TBLISTA_Contas_Recebidas.AbsolutePage)
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
    TBLISTA_Contas_Recebidas.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Contas_Recebidas.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Recebidas.AbsolutePage = 1
ProcExibePagina (TBLISTA_Contas_Recebidas.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Recebidas.AbsolutePage <> -3 Then
  If TBLISTA_Contas_Recebidas.AbsolutePage = 1 Then
    ProcExibePagina (2)
  Else
    ProcExibePagina (TBLISTA_Contas_Recebidas.AbsolutePage)
  End If
Else
   ProcExibePagina (TBLISTA_Contas_Recebidas.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Recebidas.AbsolutePage = TBLISTA_Contas_Recebidas.PageCount
ProcExibePagina (TBLISTA_Contas_Recebidas.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdproposta_Click()
On Error GoTo tratar_erro

Proposta = True
If txtProposta.Text <> "" Then
    NomeRel = "Contas_recebidas.rpt"
    ProcConstruirFiltroPadrao "PN.Proposta = '" & txtProposta & "'", "{tbl_proposta_nota.proposta} = '" & txtProposta & "'", "data_pagamento"
    ProcSalvarDadosRel False, False, False, False, Date, Date
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open StrSql_Contas_Recebidas, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        ProcConstruirFiltroPadrao "CR.proposta = '" & txtProposta & "'", "{tbl_contas_receber.proposta} = '" & txtProposta & "'", "data_pagamento"
    End If
    TBAbrir.Close
    Imprimir = True
Else
    ProcFiltrarTodas
End If
ProcCarregaLista (1)
Proposta = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdrecebimento_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "data_pagamento = '" & Format(mskData_pagamento.Value, "Short Date") & "'", "{tbl_Contas_receber.data_pagamento} = Date(" & Year(mskData_pagamento.Value) & "," & Month(mskData_pagamento.Value) & "," & Day(mskData_pagamento.Value) & ")", True, False, False, False, True, mskData_pagamento, mskData_pagamento, "data_pagamento"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro

If txtStatus.Text <> "" Then
    If txtStatus = "DUPLICATA DESCONTADA RECOMPRADA" Then
        NomeRel = "Contas_recebidas_recomprada.rpt"
    ElseIf txtStatus = "DUPLICATA DESCONTADA EM ABERTO" Then
            NomeRel = "Contas_recebidas_statusdescontada.rpt"
    End If
    StrSql_Contas_Recebidas = "Select * from tbl_Contas_RECEBER WHERE status = '" & txtStatus.Text & "' and logsit= 'S' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by tbl_Contas_receber.data_pagamento desc, tbl_Contas_receber.IdIntConta"
    NomeRel = "Contas_recebidas.rpt"
    FormulaRel_Contas_Recebidas = "{tbl_Contas_receber.status} = '" & txtStatus & "' and {tbl_contas_receber.LogSit} = 'S' and {tbl_contas_receber.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    ProcSalvarDadosRel False, False, False, False, Date, Date
    Imprimir = True
Else
    ProcFiltrarTodas
End If
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdtipo_Click()
On Error GoTo tratar_erro

If cmbtipo_conta.Text <> "" Then
    ProcFiltrarContas "Tipo_doc = '" & cmbtipo_conta.Text & "'", "{tbl_Contas_receber.Tipo_doc} = '" & cmbtipo_conta.Text & "'", True, False, False, False, False, Date, Date, "data_pagamento"
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

ProcFiltrarContas "vencimento = '" & Format(mskVencimento.Value, "Short Date") & "'", "{tbl_Contas_receber.vencimento} = Date(" & Year(mskVencimento.Value) & "," & Month(mskVencimento.Value) & "," & Day(mskVencimento.Value) & ")", True, False, False, True, False, mskVencimento, mskVencimento, "data_pagamento"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF6: ProcPlanoContas
    Case vbKeyF7: ProcCancelarRecompra
    Case vbKeyF8: ProcAntecipacoes
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

ProcCarregaToolBar1 Me, 15195, 11, True

Formulario = "Financeiro/Contas recebidas"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False
Imprimir = False
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
Txt_data_transacao.Value = Date
mskEmissao.Value = Date
mskVencimento.Value = Date
mskData_pagamento.Value = Date
Cmb_data_movimentacao.Value = Date
Cmb_tipo = "Cliente"
Cmb_opcao_lista = "Excluir"
ProcCarregaComboAno cmbAno, "2005", 1
TabFiltro.Tabs(Month(Date)).Selected = True

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboBanco()
On Error GoTo tratar_erro

ProcCarregaComboBancoFinanceiro txtBanco, "txt_Descricao <> 'Null' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

ProcCarregaComboFormaPgtoRcbto txtFormaPagto, "Tipo = 'R'"
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Contas recebidas"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
NomeRel = "Contas_recebidas.rpt"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Imprimir = True Then
frmContas_recebidas_menuimpressao.Show 1
Else
USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
End If

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
                If USMsgBox("Deseja realmente cancelar a baixa dessa(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Parcial = True Then
                    Qtd_Prog = 0
                    If TBContas!Antecipacao = False Then ProcAtualizaSaldoAntecipacao .ListItems(InitFor)
                    If TBContas!Devolucao = True Then procExcluirDevolucao .ListItems(InitFor)
                    
                    'Verifica valor do fluxo (normal ou com antecipação)
                    Set TBFluxo = CreateObject("adodb.recordset")
                    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFluxo.EOF = False Then
                        Valor3 = TBFluxo!valor
                    End If
                    
                    'Fluxo de caixa
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contas_receber where IdIntConta = " & TBContas!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo)
                        TBCorretiva.Delete
                    End If
                    TBCorretiva.Close
                    
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contas_receber where IdIntConta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        ValorParcial = TBContas!valortitulorecebido
                        Pendente = TBCorretiva!valor
                        TBCorretiva!valor = (Pendente + ValorParcial)
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_receber where tituloref = '" & IIf(TBContas!tituloref = "", 0, TBContas!tituloref) & "' and IdIntConta <> " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then
                                If TBCorretiva!titulodesc = False Then TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL" Else TBCorretiva!status = "DUPLICATA DESCONTADA EM ABERTO"
                            End If
                        Else
                            If TBCorretiva!Bloqueado = False Then
                                If TBCorretiva!titulodesc = False Then TBCorretiva!status = "TÍTULO EM ABERTO" Else TBCorretiva!status = "DUPLICATA DESCONTADA EM ABERTO"
                            End If
                            TBCorretiva!Parcial = False
                            TBCorretiva!RecebidoParcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                        
                        'Fluxo de Caixa
                        If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Or TBContas!FormaBaixa = "DOC" Or TBContas!FormaBaixa = "TED" Or TBContas!FormaBaixa = "MALOTE" Then
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
                            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where " & TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                TBFluxo!valor = Format(TBFluxo!valor - Valor3, "###,##0.00")
                                TBFluxo.Update
                                If TBFluxo!valor <= 0 Then TBFluxo.Delete
                            End If
                        End If
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                        TBFluxo!Operacao = "À Creditar"
                        TBFluxo!Data = TBCorretiva!Vencimento
                        TBFluxo!valor = TBCorretiva!valor
                        TBFluxo!Descricao = TBCorretiva!Nome_Razao
                        TBFluxo!status = "N"
                        TBFluxo!int_NotaFiscal = TBContas!NFiscal
                        TBFluxo!Documento = TBContas!txt_ndocumento
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!Cheque = 0
                        If TBCorretiva!titulodesc = False Then TBFluxo!Bloqueado = False
                        TBFluxo.Update
                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                        TBFluxo.Close
                        
                        If TBCorretiva!titulodesc = False And TBContas!FormaBaixa <> "CHEQUE" And TBContas!FormaBaixa <> "CHEQUE PRÉ-DATADO" Then
                            'Verifica saldo da antecipação
                            Qtd = .ListItems(InitFor).SubItems(8)
                            Qtde = 0
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "SELECT Sum(Valor) as valor from tbl_Contas_antecipacao where tbl_contas_antecipacao.id_conta= " & .ListItems(InitFor) & " and tbl_contas_antecipacao.tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                            TBAbrir.Close

                            qt = (Qtd - Qtde) - Qtd_Prog
                            
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select * from tbl_instituicoes where txt_descricao = '" & TBContas!Banco & "' and ID_empresa = " & TBContas!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                TBProduto!Saldo = TBProduto!Saldo - qt
                                TBProduto.Update
                            End If
                            TBProduto.Close
                        End If
                    End If
                    
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from familia_financeiro where idconta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref) & " and tipoconta = 'R' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        Do While TBFamilia.EOF = False
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from familia_financeiro where IDConta = " & TBContas!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                TBFamilia!valor = TBFamilia!valor + TBCiclo!valor
                                TBFamilia.Update
                            End If
                            TBCiclo.Close
                            
                            Conexao.Execute "DELETE from familia_financeiro where IDConta = " & TBContas!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'R' and Deposito_transf = 'False'"
                            TBFamilia.MoveNext
                        Loop
                    End If
                    TBFamilia.Close
                    TBCorretiva.Update
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select local_troca from Troca_titulo where ID = " & TBCorretiva!IDtrocatitulo, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Set TBReceber = CreateObject("adodb.recordset")
                        TBReceber.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBFI!local_troca & "' and tbl_contas_receber.ID_empresa = " & TBCorretiva!ID_empresa & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBReceber.EOF = False Then
                            valor = IIf(IsNull(TBReceber!valor), 0, TBReceber!valor)
                            NovoValor = Replace(valor, ",", ".")
                            Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & TBFI!local_troca & "' and ID_empresa = " & TBCorretiva!ID_empresa
                        End If
                        TBReceber.Close
                    End If
                    TBFI.Close
                    
                    TBCorretiva.Close
                Else
                    If TBContas!Antecipacao = False Then ProcAtualizaSaldoAntecipacao .ListItems(InitFor)
                    If TBContas!Devolucao = True Then procExcluirDevolucao .ListItems(InitFor)
                    
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contas_receber where IdIntConta = " & TBContas!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_receber where tituloref = '" & TBContas!IDintconta & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then
                                If TBCorretiva!titulodesc = False Then TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL" Else TBCorretiva!status = "DUPLICATA DESCONTADA EM ABERTO"
                            End If
                        Else
                            If TBCorretiva!Bloqueado = False Then
                                If TBCorretiva!titulodesc = False Then TBCorretiva!status = "TÍTULO EM ABERTO" Else TBCorretiva!status = "DUPLICATA DESCONTADA EM ABERTO"
                            End If
                            TBCorretiva!Parcial = False
                            TBCorretiva!RecebidoParcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                        
                        'Verifica valor do fluxo (normal ou com antecipação)
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = False Then
                            Valor3 = TBFluxo!valor
                        End If
                        
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
                                TBFluxo!valor = Format(TBFluxo!valor - Valor3, "###,##0.00")
                                TBFluxo.Update
                                If TBFluxo!valor <= 0 Then
                                    TBFluxo.Delete
                                    Conexao.Execute "DELETE from tbl_Contas_Varias where ID = " & IIf(IsNull(TBContas!ID_varias), 0, TBContas!ID_varias)
                                End If
                            End If
                        End If
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                        TBFluxo!Operacao = "À Creditar"
                        TBFluxo!Data = TBCorretiva!Vencimento
                        TBFluxo!valor = TBCorretiva!valor
                        TBFluxo!Descricao = TBCorretiva!Nome_Razao
                        TBFluxo!status = "N"
                        TBFluxo!int_NotaFiscal = TBContas!NFiscal
                        TBFluxo!Documento = TBContas!txt_ndocumento
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!Cheque = 0
                        If TBCorretiva!titulodesc = False Then TBFluxo!Bloqueado = False
                        TBFluxo.Update
                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                        TBFluxo.Close
                        
                        TBCorretiva!NDoctoBaixa = ""
                        TBCorretiva!Obs = ""
                        TBCorretiva!Logsit = "N"
                        TBCorretiva!Data_pagamento = Null
                        TBCorretiva!Data_movimentacao = Null
                        TBCorretiva!diferenca = 0
                        TBCorretiva!valor_tituloUSS = 0
                        TBCorretiva!valorUSSemicao = 0
                        TBCorretiva!valorUSSrecebimento = 0
                        TBCorretiva!valortitulorecebido = 0
                        TBCorretiva!exportacao = False
                        TBCorretiva!Moeda = 0
                        TBCorretiva!Dias_atraso = 0
                        TBCorretiva!Juros = 0
                        TBCorretiva!Juros_valor = 0
                        TBCorretiva!Multa = 0
                        TBCorretiva!Multa_valor = 0
                        TBCorretiva!Desconto = 0
                        TBCorretiva!Desconto_valor = 0
                    End If
                    
                    If TBCorretiva!titulodesc = False And TBContas!FormaBaixa <> "CHEQUE" And TBContas!FormaBaixa <> "CHEQUE PRÉ-DATADO" Then
                        'Verifica saldo da antecipação
                        Qtd = .ListItems(InitFor).SubItems(8)
                        Qtde = 0
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "SELECT Sum(Valor) as valor from tbl_Contas_antecipacao where id_conta = " & .ListItems(InitFor) & " and tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                        TBAbrir.Close
                        
                        qt = Qtd - Qtde
                        
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select * from tbl_instituicoes where txt_descricao = '" & TBContas!Banco & "' and ID_empresa = " & TBContas!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then
                            If TBContas!Devolucao = True Then TBProduto!Saldo = Format(TBProduto!Saldo + Qtd_Prog, "###,##0.00") Else TBProduto!Saldo = Format(TBProduto!Saldo - qt, "###,##0.00")
                            TBProduto.Update
                        End If
                        TBProduto.Close
                    End If
                    
                    Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where idconta = " & TBContas!IDintconta & " and tipoconta = 'R'"
                    
                    TBCorretiva!ID_varias = 0
                    TBCorretiva.Update
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select local_troca from Troca_titulo where ID = " & IIf(IsNull(TBCorretiva!IDtrocatitulo), 0, TBCorretiva!IDtrocatitulo), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Set TBReceber = CreateObject("adodb.recordset")
                        TBReceber.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBFI!local_troca & "' and tbl_contas_receber.ID_empresa = " & TBCorretiva!ID_empresa & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBReceber.EOF = False Then
                            valor = IIf(IsNull(TBReceber!valor), 0, TBReceber!valor)
                            NovoValor = Replace(valor, ",", ".")
                            Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & TBFI!local_troca & "' and ID_empresa = " & TBCorretiva!ID_empresa
                        End If
                        TBReceber.Close
                    End If
                    TBFI.Close
                    
                    TBCorretiva.Close
                End If
                Conexao.Execute "DELETE from tbl_contas_antecipacao where ID_Conta = " & .ListItems(InitFor) & " and tipo = 'R'"
                
                '==================================
                Modulo = "Financeiro/Contas recebidas"
                ID_documento = .ListItems(InitFor)
                Evento = "Cancelar baixa"
                Documento = "Documento: " & TBContas!txt_ndocumento
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBContas.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) conta(s) antes de cancelar a baixa."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Baixa(s) cancelada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Lista.SetFocus
    ProcCarregaDados
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldoAntecipacao(IDConta As Long)
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select CA.* FROM tbl_Contas_receber CR INNER JOIN tbl_contas_antecipacao CA ON CR.IdIntconta = CA.ID_conta where CR.IdIntConta = " & IDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Do While TBFI.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Saldo_antecipacao, LogSit from tbl_Contas_receber where IDintconta = " & TBFI!ID_antecipacao, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir!Saldo_antecipacao = TBAbrir!Saldo_antecipacao + TBFI!valor
            If TBAbrir!Saldo_antecipacao = 0 Then TBAbrir!Logsit = "S" Else TBAbrir!Logsit = "N"
            TBAbrir.Update
        End If
        TBAbrir.Close
        TBFI.MoveNext
    Loop
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmContas_recebidas_localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Imprimir = False
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
If txtidintconta.Text = "" Then
    USMsgBox ("Informe a conta antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If txtidintconta.Text = "" Then
    NomeCampo = "a conta"
    ProcVerificaAcao
    Exit Sub
End If
If txtFormaPagto = "CHEQUE" Or txtFormaPagto = "CHEQUE PRÉ-DATADO" Then
    Cheque = "Cheque n. " & txt_Ndocto
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtBanco & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False' and Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Não é permitido alterar a baixa em cheque desta conta, pois o mesmo já está compensado."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If


Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & txtidintconta.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    ID_varias = IIf(IsNull(TBContas!ID_varias), 0, TBContas!ID_varias)
    If (txtFormaPagto = "CHEQUE" Or txtFormaPagto = "CHEQUE PRÉ-DATADO" Or txtFormaPagto = "DOC" Or txtFormaPagto = "TED" Or txtFormaPagto = "MALOTE" Or ID_varias <> 0) And TBContas!status <> "DUPLICATA DESCONTADA EM ABERTO" Then
        TextoFiltroData = "Data = '" & Format(TBContas!Data_movimentacao, "Short Date") & "' and"
        Select Case txtFormaPagto
            Case "CHEQUE":
                Descricao = "Cheque n. " & txt_Ndocto
                TextoFiltroData = ""
            Case "CHEQUE PRÉ-DATADO":
                Descricao = "Cheque n. " & txt_Ndocto
                TextoFiltroData = ""
            Case "DOC": Descricao = "Doc n. " & txt_Ndocto
            Case "TED": Descricao = "Ted n. " & txt_Ndocto
            Case "MALOTE": Descricao = "Malote n. " & txt_Ndocto
        End Select
        
        If ID_varias = 0 Then
            TextoFiltro1 = TextoFiltroData & " Operacao = 'Crédito' and Descricao = '" & Descricao & "' and Cheque = '" & txt_Ndocto & "' and Instituicao = '" & cmb_Banco & "'"
        Else
            TextoFiltro1 = "ID_varias = " & ID_varias
        End If
        
        Set TBFluxo = CreateObject("adodb.recordset")
        TBFluxo.Open "Select Data from tbl_Fluxo_de_caixa where " & TextoFiltro1, Conexao, adOpenKeyset, adLockOptimistic
        If TBFluxo.EOF = False Then
            TBFluxo!Data = Cmb_data_movimentacao
            TBFluxo.Update
        End If
        TBFluxo.Close
    End If
    
    TBContas!Data_movimentacao = Cmb_data_movimentacao
    TBContas!Observacoes = txtObservacao
    TBContas!Obs = txtobservacao_recbto
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select Data from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        TBFluxo!Data = Cmb_data_movimentacao
        TBFluxo.Update
    End If
        
    If ID_varias <> 0 Then Conexao.Execute "UPDATE tbl_contas_receber Set Data_movimentacao = '" & Cmb_data_movimentacao & "' where IDintConta <> " & txtidintconta & " and ID_varias = " & ID_varias
    
    TBContas.Update
End If
TBContas.Close
USMsgBox ("Alteração da data da movimentação e observações efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Financeiro/Contas recebidas"
ID_documento = txtidintconta
Evento = "Alterar"
Documento = "Documento: " & txtDocumento
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    Lista.SetFocus
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
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
                    Set TBContas = CreateObject("adodb.recordset")
                    TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBContas.EOF = False Then
                        'Verifica se a conta parcial já está líquidada
                        If TBContas!Parcial = True Then
                            Set TBCorretiva = CreateObject("adodb.recordset")
                            TBCorretiva.Open "Select * from tbl_contas_receber where IdIntConta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                            If TBCorretiva.EOF = False Then
                                If TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL LIQUIDADO" And TBContas!status <> "TÍTULO RECEBIDO PARCIAL LIQUIDADO" Then GoTo Proximo
                            End If
                            TBCorretiva.Close
                        End If
                        
                        If TBContas!status = "DUPLICATA DESCONTADA RECOMPRADA" Then GoTo Proximo
                        
                        If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                            If IsNull(TBContas!NDoctoBaixa) = False And TBContas!NDoctoBaixa <> "" Then
                                Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False' and Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = False Then GoTo Proximo
                                TBAbrir.Close
                            End If
                        End If
                        If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                        
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_instituicoes_transf where IdIntConta = " & .ListItems(InitFor) & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then GoTo Proximo
                        TBAbrir.Close
                        
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_devolucao where ID_Conta = '" & .ListItems(InitFor) & "' and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            TBContas.Close
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                        TBAbrir.Close
                        
                    End If
                    TBContas.Close
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

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Cmb_opcao_lista = "Excluir" Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    'Verifica se a conta parcial já está líquidada
                    If TBContas!Parcial = True Then
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select * from tbl_contas_receber where IdIntConta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            If TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL LIQUIDADO" And TBContas!status <> "TÍTULO RECEBIDO PARCIAL LIQUIDADO" Then
                                USMsgBox ("Não é permitido cancelar a baixa desta conta, pois ela já está líquidada."), vbExclamation, "CAPRIND v5.0"
                                TBContas.Close
                                TBCorretiva.Close
                                .ListItems.Item(InitFor).Checked = False
                                Exit Sub
                            End If
                        End If
                        TBCorretiva.Close
                    End If
                    
                    If TBContas!status = "DUPLICATA DESCONTADA RECOMPRADA" Then
                        USMsgBox ("Não é permitido cancelar a baixa desta conta, pois a mesma está recomprada."), vbExclamation, "CAPRIND v5.0"
                        TBContas.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    
                    If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                        If IsNull(TBContas!NDoctoBaixa) = False And TBContas!NDoctoBaixa <> "" Then
                            Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False' and Operacao = 'Crédito'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                USMsgBox ("Não é permitido cancelar a baixa em cheque desta conta, pois o mesmo já está compensado."), vbExclamation, "CAPRIND v5.0"
                                TBContas.Close
                                TBAbrir.Close
                                .ListItems.Item(InitFor).Checked = False
                                Exit Sub
                            End If
                            TBAbrir.Close
                        End If
                    End If
                    
                    If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                        USMsgBox ("Não é permitido cancelar a baixa desta conta, pois a mesma é uma antecipação liquídada."), vbExclamation, "CAPRIND v5.0"
                        TBContas.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_instituicoes_transf where IdIntConta = " & .ListItems(InitFor) & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido cancelar a baixa desta conta por este módulo, pois a mesma é uma tarifa bancária."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                    TBAbrir.Close
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_contas_devolucao where ID_Conta = '" & .ListItems(InitFor) & "' and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido cancelar a baixa desta conta, pois a mesma está vinculada a uma devolução."), vbExclamation, "CAPRIND v5.0"
                        TBContas.Close
                        TBAbrir.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    TBAbrir.Close
                    
                End If
                TBContas.Close
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
ProcCarregaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

ProcLimpaCampos
If Lista.ListItems.Count = 0 Then Exit Sub
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    If IsNull(TBContas!ID_empresa) = False And TBContas!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBContas!ID_empresa
    txtidintconta.Text = TBContas!IDintconta
    Txt_data_transacao.Value = IIf(IsNull(TBContas!Data_transacao), Date, Format(TBContas!Data_transacao, "dd/mm/yyyy"))
    txtDocumento = IIf(IsNull(TBContas!txt_ndocumento), "", TBContas!txt_ndocumento)
    txtNFiscal.Text = IIf(IsNull(TBContas!NFiscal), "", TBContas!NFiscal)
    mskEmissao.Value = TBContas!emissao
    
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
    txtNome_Razao.Text = TBContas!Nome_Razao
    txtCidade.Text = IIf(IsNull(TBContas!Cidade), "", TBContas!Cidade)
    txtuf.Text = IIf(IsNull(TBContas!Estado), "", TBContas!Estado)
    txtObservacao.Text = IIf(IsNull(TBContas!Observacoes), "", TBContas!Observacoes)
    mskVencimento.Value = TBContas!Vencimento
    
    If IsNull(TBContas!Parcela) = False And TBContas!Parcela <> "" Then txtparcela.Text = TBContas!Parcela
    
    'Dados de Recebimento
    LblDocumento.Caption = "N° documento baixa"
    Select Case txtFormaPagto
        Case "DOC": LblDocumento.Caption = "N° do DOC"
        Case "TED": LblDocumento.Caption = "N° do TED"
        Case "CHEQUE": LblDocumento.Caption = "N° do cheque"
        Case "CHEQUE PRÉ-DATADO": LblDocumento.Caption = "N° do cheque"
        Case "MALOTE": LblDocumento.Caption = "N° do malote"
    End Select
    
    mskData_pagamento.Value = IIf(IsNull(TBContas!Data_pagamento), Date, Format(TBContas!Data_pagamento, "dd/mm/yyyy"))
    Cmb_data_movimentacao.Value = IIf(IsNull(TBContas!Data_movimentacao), Date, Format(TBContas!Data_movimentacao, "dd/mm/yyyy"))
    txtobservacao_recbto.Text = IIf(IsNull(TBContas!Obs), "", TBContas!Obs)
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Local_troca from troca_titulo where ID = " & IIf(IsNull(TBContas!IDtrocatitulo), 0, TBContas!IDtrocatitulo), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_local_desconto = IIf(IsNull(TBAbrir!local_troca), "", TBAbrir!local_troca)
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select txt_Conta from tbl_Instituicoes where txt_Descricao = '" & TBContas!Banco & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtConta.Text = IIf(IsNull(TBAbrir!txt_Conta), "", TBAbrir!txt_Conta)
    End If
    If TBContas!Parcial = True And txtStatus.Text <> "TÍTULO RECEBIDO PARCIAL LIQUIDADO" Then
        chbparcial.Value = 1
        txtValor = Format(IIf(IsNull(TBContas!RecebidoParcial), 0, TBContas!RecebidoParcial) + IIf(IsNull(TBContas!ValorPendente), 0, TBContas!ValorPendente), "###,##0.00")
    Else
        chbparcial.Value = 0
        txtValor = IIf(IsNull(TBContas!valor), "", Format(TBContas!valor, "###,##0.00"))
    End If
    txtvalortitrecebido = IIf(IsNull(TBContas!valortitulorecebido), "", Format(TBContas!valortitulorecebido, "###,##0.00"))
    Txt_dias_atraso = IIf(IsNull(TBContas!Dias_atraso), "", TBContas!Dias_atraso)
    txtjuros.Text = IIf(IsNull(TBContas!Juros_valor), 0, Format(TBContas!Juros_valor, "###,##0.0000000"))
    Txt_total_juros = Format(IIf(IsNull(TBContas!Juros_valor), 0, TBContas!Juros_valor) * IIf(IsNull(TBContas!Dias_atraso), 0, TBContas!Dias_atraso), "###,##0.0000000")
    Txt_multa = IIf(IsNull(TBContas!Multa_valor), 0, Format(TBContas!Multa_valor, "###,##0.0000000"))
    txtDesconto.Text = IIf(IsNull(TBContas!Desconto_valor), 0, Format(TBContas!Desconto_valor, "###,##0.0000000"))
    txt_Ndocto.Text = IIf(IsNull(TBContas!NDoctoBaixa), "", TBContas!NDoctoBaixa)
    txt_tituloref.Text = IIf(IsNull(TBContas!tituloref), "", TBContas!tituloref)
    ProcCarregaProposta
    CodigoLista = Lista.SelectedItem.index
    
    NomeCampo = "o tipo do documento"
    If IsNull(TBContas!Tipo_doc) = False And TBContas!Tipo_doc <> "" Then cmbtipo_conta.Text = TBContas!Tipo_doc
    NomeCampo = "o status"
    If IsNull(TBContas!status) = False And TBContas!status <> "" Then txtStatus.Text = TBContas!status
    NomeCampo = "o banco"
    If IsNull(TBContas!Banco) = False And TBContas!Banco <> "" Then txtBanco.Text = TBContas!Banco
    NomeCampo = "a forma da baixa"
    If IsNull(TBContas!FormaBaixa) = False And TBContas!FormaBaixa <> "" Then txtFormaPagto.Text = TBContas!FormaBaixa
1:
        ProcCarregaListaPC
End If
TBContas.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaPC()
On Error GoTo tratar_erro

Lista_PC.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select FF.ID, F.Codigo, F.txt_descricao, FF.Valor from Familia_financeiro FF INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDConta = " & txtidintconta & " and FF.Tipoconta = 'R' and FF.Pago_recebido = 'True' and FF.Deposito_transf = 'False' order by F.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_PC.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
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
    TBAbrir.Open "Select Valortitulorecebido, Devolucao from tbl_contas_receber where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qt = TBAbrir!valortitulorecebido
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
    Modulo = "Financeiro/Contas recebidas"
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

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

ProcCorrigeLayoutForm

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeLayoutForm()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Frame2.Top = Frame3.Top + Frame3.Height + 10
        With Lista
            .Top = Frame2.Top + Frame2.Height + 10
            .Height = Frame1.Top - .Top
            If .Visible = True Then .SetFocus
        End With
    Case 1:
        Frame2.Top = Frame4.Top + Frame4.Height
        With Lista
            .Top = Frame2.Top + Frame2.Height
            .Height = Frame1.Top - .Top
            If .Visible = True Then .SetFocus
        End With
End Select

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

NomeRel = "Contas_recebidas.rpt"
If TabFiltro.SelectedItem.key = "Todas" Then
    ProcFiltrarTodas
Else
    M = FunVerificaMes(TabFiltro.SelectedItem.key)
    If OptDomes.Value = True Then ProcConstruirFiltroPadrao "month(CR.data_pagamento)= '" & M & "' and Year(CR.data_pagamento) = '" & cmbAno & "'", "Month ({tbl_contas_receber.data_pagamento}) = " & M & " and year ({tbl_Contas_receber.data_pagamento})= " & cmbAno, "data_pagamento"
    If OptAteomes.Value = True Then ProcConstruirFiltroPadrao "month(CR.data_pagamento)<= '" & M & "' and Year(CR.data_pagamento) = '" & cmbAno & "'", "Month ({tbl_Contas_receber.data_pagamento}) <= " & M & " and year ({tbl_Contas_receber.data_pagamento})= " & cmbAno, "data_pagamento"
    ProcSalvarDadosRel False, False, False, False, Date, Date
    ProcCarregaLista (1)
    Imprimir = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtBanco_Click()
On Error GoTo tratar_erro
    
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select txt_conta from tbl_instituicoes where txt_descricao = '" & txtBanco & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtConta.Text = IIf(IsNull(TBAbrir!txt_Conta), "", TBAbrir!txt_Conta)
End If
TBAbrir.Close

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
        TBAbrir.Open "Select NomeRazao, Cidade, UF from Clientes where idcliente = " & txtIDcliente & " and Prospecto = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtNome_Razao = IIf(IsNull(TBAbrir!NomeRazao), "", TBAbrir!NomeRazao)
            txtCidade = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
            cbo_UF = IIf(IsNull(TBAbrir!UF), "", TBAbrir!UF)
        End If
    ElseIf Cmb_tipo = "Fornecedor" Then
            TBAbrir.Open "Select Nome_Razao, Cidade, Estado from compras_fornecedores where idcliente = " & txtIDcliente & " and Prospecto = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                txtNome_Razao = IIf(IsNull(TBAbrir!Nome_Razao), "", TBAbrir!Nome_Razao)
                txtCidade = IIf(IsNull(TBAbrir!Cidade), "", TBAbrir!Cidade)
                cbo_UF = IIf(IsNull(TBAbrir!Estado), "", TBAbrir!Estado)
            End If
        ElseIf Cmb_tipo = "Funcionário" Then
                TBAbrir.Open "Select Nome from Funcionarios where Codigo = '" & txtIDcliente & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtNome_Razao = TBAbrir!Nome
            Else
                TBAbrir.Open "Select Txt_descricao from tbl_Instituicoes where ID = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtNome_Razao = TBAbrir!Txt_descricao
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
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

Private Sub txtvalortitrecebido_LostFocus()
On Error GoTo tratar_erro

If txtvalortitrecebido.Text <> "" Then
    VerifNumero = txtvalortitrecebido.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvalortitrecebido.Text = ""
        txtvalortitrecebido.SetFocus
        Exit Sub
    End If
    txtvalortitrecebido.Text = Format(txtvalortitrecebido.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcImprimir
    Case 5: ProcPlanoContas
    Case 6: ProcCancelarRecompra
    Case 7: ProcAntecipacoes
    Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procExcluirDevolucao(IDConta As Long)
On Error GoTo tratar_erro

Qtd_Prog = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(valor) as valor from tbl_contas_devolucao where Id_devolucao = " & IDConta & " and tipo = 'R' and logsit = 'S'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Qtd_Prog = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
TBAbrir.Close

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_contas_devolucao where ID_Devolucao = " & IDConta & " and Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
Do While TBFI.EOF = False
    If TBFI!Logsit = "N" Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from tbl_Contas_receber where IdIntConta = " & TBFI!ID_conta, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            If TBItem!Parcial = False Then
                If TBItem!titulodesc = False Then
                    If TBItem!Bloqueado = False Then TBItem!status = "TÍTULO EM ABERTO"
                Else
                    TBItem!status = "DUPLICATA DESCONTADA EM ABERTO"
                End If
                TBItem!Parcial = False
                TBItem!RecebidoParcial = 0
                TBItem!ValorPendente = 0
                TBItem!tituloref = ""
                TBItem!valorprincipal = 0
                TBItem!Logsit = "N"
                TBItem!Data_pagamento = Null
                TBItem!Data_movimentacao = Null
                TBItem!valortitulorecebido = 0
                TBItem!NDoctoBaixa = ""
                TBItem!Obs = ""
                TBItem!diferenca = 0
                TBItem!valor_tituloUSS = 0
                TBItem!valorUSSemicao = 0
                TBItem!valorUSSrecebimento = 0
                TBItem!exportacao = False
                TBItem!Moeda = 0
                TBItem!Dias_atraso = 0
                TBItem!Juros = 0
                TBItem!Juros_valor = 0
                TBItem!Multa = 0
                TBItem!Multa_valor = 0
                TBItem!Desconto = 0
                TBItem!Desconto_valor = 0
                TBItem!ID_varias = 0
                TBItem.Update
                Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where idconta = " & TBFI!ID_conta & " and tipoconta = 'R'"
            Else
                Conexao.Execute "DELETE from F from tbl_Fluxo_de_caixa F INNER JOIN tbl_Contas_receber CR ON CR.IDFluxo = F.IDFluxo where CR.IdIntConta = " & IIf(IsNull(TBFI!ID_conta), 0, TBFI!ID_conta)
                Conexao.Execute "DELETE from tbl_Contas_receber where IdIntConta = " & IIf(IsNull(TBFI!ID_conta), 0, TBFI!ID_conta)
                
                Set TBCorretiva = CreateObject("adodb.recordset")
                TBCorretiva.Open "Select * from tbl_contas_receber where IdIntConta = " & IIf(TBItem!tituloref = "", 0, TBItem!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                If TBCorretiva.EOF = False Then
                    ValorParcial = TBItem!valortitulorecebido
                    Pendente = TBItem!ValorPendente
                    If TBCorretiva!Logsit = "N" Then
                        TBCorretiva!valor = (Pendente + ValorParcial)
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_receber where tituloref = '" & IIf(TBItem!tituloref = "", 0, TBItem!tituloref) & "' and IdIntConta <> " & TBItem!tituloref, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL"
                        Else
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                            TBCorretiva!Parcial = False
                            TBCorretiva!RecebidoParcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                        
                        Set TBFamilia = CreateObject("adodb.recordset")
                        TBFamilia.Open "Select * from familia_financeiro where idconta = " & IIf(TBItem!tituloref = "", 0, TBItem!tituloref) & " and tipoconta = 'R' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFamilia.EOF = False Then
                            Do While TBFamilia.EOF = False
                                Set TBCiclo = CreateObject("adodb.recordset")
                                TBCiclo.Open "Select * from familia_financeiro where IDConta = " & TBItem!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBCiclo.EOF = False Then
                                    TBFamilia!valor = TBFamilia!valor + TBCiclo!valor
                                    TBFamilia.Update
                                End If
                                TBCiclo.Close
                                
                                Conexao.Execute "DELETE from familia_financeiro where IDConta = " & TBItem!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'R' and Deposito_transf = 'False'"
                                TBFamilia.MoveNext
                            Loop
                        End If
                    Else
                        TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL"
                        TBCorretiva!Parcial = True
                        TBCorretiva!RecebidoParcial = TBCorretiva!valor
                        TBCorretiva!ValorPendente = Format(TBCorretiva!valorprincipal - TBCorretiva!valortitulorecebido, "###,##0.00")
                        TBCorretiva!tituloref = ""
                    End If
                    TBCorretiva.Update
                End If
                TBCorretiva.Close
            End If
        End If
        TBItem.Close
    End If
    TBFI.MoveNext
Loop
Conexao.Execute "DELETE from tbl_contas_devolucao where ID_Devolucao = " & IDConta & " and Tipo = 'R'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarDadosRel(DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataBaixa As Boolean, DataInicio As Date, DataFinal As Date)
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = Formulario
    If DataTransacao = True Or DataEmissao = True Or DataVencimento = True Or DataBaixa = True Then
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
