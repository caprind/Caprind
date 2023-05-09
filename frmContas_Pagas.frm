VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContas_Pagas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Contas pagas"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmContas_Pagas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      Left            =   2940
      MaxLength       =   20
      TabIndex        =   103
      Top             =   5580
      Visible         =   0   'False
      Width           =   1935
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      ToolTipText     =   "Número da conta."
      Top             =   5580
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   95
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
         ItemData        =   "frmContas_Pagas.frx":1042
         Left            =   6960
         List            =   "frmContas_Pagas.frx":104C
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
         DibPicture      =   "frmContas_Pagas.frx":1064
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
         DibPicture      =   "frmContas_Pagas.frx":4808
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
         DibPicture      =   "frmContas_Pagas.frx":8311
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
         DibPicture      =   "frmContas_Pagas.frx":C400
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
         TabIndex        =   108
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
         TabIndex        =   106
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
         TabIndex        =   97
         Top             =   240
         Width           =   1095
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
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   75
      TabIndex        =   82
      Top             =   9390
      Width           =   15195
      Begin VB.TextBox TotalContas 
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
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Total geral pago."
         Top             =   180
         Width           =   1620
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   94
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
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12270
         TabIndex        =   83
         Top             =   180
         Width           =   2445
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      ButtonCaption6  =   "Centro de custo"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Centro de custo (F7)"
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
      ButtonWidth6    =   97
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Visualizar"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Visualizar relacionamento (F8)"
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
      ButtonLeft7     =   368
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
      ButtonLeft8     =   432
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
      ButtonLeft9     =   436
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
      ButtonLeft10    =   479
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
      ButtonLeft11    =   511
      ButtonTop11     =   2
      ButtonWidth11   =   24
      ButtonHeight11  =   24
      ButtonUseMaskColor11=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   9990
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmContas_Pagas.frx":FC8C
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView lst_ContasPagas 
      Height          =   4065
      Left            =   75
      TabIndex        =   29
      Top             =   4695
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7170
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. vencto."
         Object.Width           =   1940
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
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   6094
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. baixado"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "D"
         Text            =   "Dt. baixa"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame4 
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
      Top             =   4155
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
         ItemData        =   "frmContas_Pagas.frx":15F06
         Left            =   14220
         List            =   "frmContas_Pagas.frx":15F08
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   63
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
         TabIndex        =   60
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
         TabIndex        =   61
         Top             =   270
         Width           =   1035
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   405
         Left            =   2130
         TabIndex        =   62
         Top             =   240
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   714
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
      TabIndex        =   64
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
      TabPicture(0)   =   "frmContas_Pagas.frx":15F0A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados da baixa"
      TabPicture(1)   =   "frmContas_Pagas.frx":15F26
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   2055
         Left            =   -74925
         TabIndex        =   73
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14695
            Picture         =   "frmContas_Pagas.frx":15F42
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Visualizar arquivo."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmContas_Pagas.frx":16504
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Limpar caminho."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_valor_pago 
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
            Picture         =   "frmContas_Pagas.frx":16642
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Filtrar por valor pago."
            Top             =   370
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
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do comprovante."
            Top             =   1590
            Width           =   6705
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14040
            Picture         =   "frmContas_Pagas.frx":16A5D
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Localizar comprovante."
            Top             =   1590
            Width           =   315
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
            ItemData        =   "frmContas_Pagas.frx":16B5F
            Left            =   10995
            List            =   "frmContas_Pagas.frx":16B99
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   49
            ToolTipText     =   "Forma da baixa."
            Top             =   370
            Width           =   4020
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
         Begin VB.TextBox txt_ValorPago 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Valor baixado."
            Top             =   370
            Width           =   1200
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
         Begin VB.TextBox txtobs_pgto 
            BackColor       =   &H00FFFFFF&
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
            Height          =   915
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            ToolTipText     =   "Observações do pagamento."
            Top             =   990
            Width           =   7065
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
         Begin VB.CommandButton cmdbaixa 
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
            Picture         =   "frmContas_Pagas.frx":16CA7
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Filtrar por data de pagamento."
            Top             =   370
            Width           =   315
         End
         Begin VB.CommandButton cmdbanco 
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
            Picture         =   "frmContas_Pagas.frx":170C2
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Filtrar por instituição bancária."
            Top             =   990
            Width           =   315
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
         Begin MSComCtl2.DTPicker txtBaixado 
            Height          =   315
            Left            =   1815
            TabIndex        =   41
            ToolTipText     =   "Data da baixa."
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
            Format          =   182517763
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
            Format          =   182517761
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
            TabIndex        =   107
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho do comprovante"
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
            Left            =   9757
            TabIndex        =   89
            Top             =   1380
            Width           =   1830
         End
         Begin VB.Label Label23 
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
            TabIndex        =   87
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label25 
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
            TabIndex        =   86
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label8 
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
            Left            =   285
            TabIndex        =   85
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label20 
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
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Observações do pagamento"
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
            Height          =   240
            Left            =   2685
            TabIndex        =   81
            Top             =   780
            Width           =   2055
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
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
            TabIndex        =   80
            Top             =   780
            Width           =   435
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
            Left            =   13510
            TabIndex        =   79
            Top             =   780
            Width           =   1505
         End
         Begin VB.Label Label5 
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
            Left            =   12173
            TabIndex        =   78
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2130
            TabIndex        =   77
            Top             =   180
            Width           =   750
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
            Index           =   1
            Left            =   12450
            TabIndex        =   76
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label19 
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
            TabIndex        =   75
            Top             =   180
            Width           =   825
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
            Index           =   0
            Left            =   10020
            TabIndex        =   74
            Top             =   180
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   2835
         Left            =   75
         TabIndex        =   65
         Top             =   1320
         Width           =   15210
         Begin VB.CheckBox chkConta_fixa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Conta fixa"
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
            Left            =   180
            TabIndex        =   26
            Top             =   1380
            Width           =   1395
         End
         Begin VB.CommandButton Cmd_localizar_contatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   10080
            Picture         =   "frmContas_Pagas.frx":174DD
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar contatos."
            Top             =   990
            Width           =   315
         End
         Begin VB.CommandButton Cmd_competencia 
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
            Left            =   11535
            Picture         =   "frmContas_Pagas.frx":177F1
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Filtrar por número do documento."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox txt_Competencia 
            BackColor       =   &H00FFFFFF&
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
            Left            =   10500
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Competência."
            Top             =   990
            Width           =   1020
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
            Left            =   1890
            Picture         =   "frmContas_Pagas.frx":17C0C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Filtrar por valor."
            Top             =   990
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
            Left            =   5985
            Picture         =   "frmContas_Pagas.frx":18027
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Filtrar por data da transação."
            Top             =   370
            Width           =   315
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
            Height          =   1095
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            ToolTipText     =   "Observações."
            Top             =   1590
            Width           =   7320
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
            ItemData        =   "frmContas_Pagas.frx":18442
            Left            =   2280
            List            =   "frmContas_Pagas.frx":18452
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Tipo."
            Top             =   990
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
            ItemData        =   "frmContas_Pagas.frx":1848E
            Left            =   180
            List            =   "frmContas_Pagas.frx":18490
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   375
            Width           =   4560
         End
         Begin VB.TextBox txt_ValorDocto 
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
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   930
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Valor."
            Top             =   990
            Width           =   960
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Número da parcela."
            Top             =   990
            Width           =   730
         End
         Begin VB.ComboBox txtNPedido 
            BackColor       =   &H00FFFFFF&
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
            Left            =   10155
            TabIndex        =   7
            ToolTipText     =   "Número do pedido de compra."
            Top             =   370
            Width           =   1280
         End
         Begin VB.CommandButton Cmdlocalizarforn 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9765
            Picture         =   "frmContas_Pagas.frx":18492
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Localizar fornecedor."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox txtIDFornec 
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
            Left            =   4185
            TabIndex        =   17
            ToolTipText     =   "Código do fornecedor."
            Top             =   990
            Width           =   810
         End
         Begin VB.CommandButton cmdpedido 
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
            Left            =   11460
            Picture         =   "frmContas_Pagas.frx":18594
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Filtrar por número do pedido de compra."
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
            Left            =   13080
            Picture         =   "frmContas_Pagas.frx":189AF
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Left            =   14700
            Picture         =   "frmContas_Pagas.frx":18DCA
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Filtrar por data de vencimento."
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
            Left            =   9765
            Picture         =   "frmContas_Pagas.frx":191E5
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Filtrar por número do documento."
            Top             =   370
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
            Left            =   7470
            Picture         =   "frmContas_Pagas.frx":19600
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Filtrar por tipo do documento."
            Top             =   370
            Width           =   315
         End
         Begin VB.TextBox txtFornec 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Nome do fornecedor."
            Top             =   990
            Width           =   4440
         End
         Begin VB.TextBox txtNDocumento 
            BackColor       =   &H00FFFFFF&
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
            Left            =   7860
            MaxLength       =   30
            TabIndex        =   5
            ToolTipText     =   "Número do documento."
            Top             =   370
            Width           =   1890
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
            ItemData        =   "frmContas_Pagas.frx":19A1B
            Left            =   6390
            List            =   "frmContas_Pagas.frx":19A4F
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Tipo do documento."
            Top             =   370
            Width           =   1065
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
            Picture         =   "frmContas_Pagas.frx":19A89
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Filtrar por status."
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
            Left            =   9450
            Picture         =   "frmContas_Pagas.frx":19EA4
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Filtrar por fornecedor."
            Top             =   990
            Width           =   315
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11940
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   990
            Width           =   2745
         End
         Begin MSComCtl2.DTPicker txtDtEmissao 
            Height          =   315
            Left            =   11850
            TabIndex        =   9
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
            Format          =   182517763
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtDataPagto 
            Height          =   315
            Left            =   13485
            TabIndex        =   11
            ToolTipText     =   "Data de vencimento."
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
            Format          =   182517763
            CurrentDate     =   39057
         End
         Begin MSComctlLib.ListView Lista_PC 
            Height          =   1095
            Left            =   7515
            TabIndex        =   28
            Top             =   1590
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
            Left            =   4755
            TabIndex        =   1
            ToolTipText     =   "Data da transação."
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
            Format          =   182517763
            CurrentDate     =   39057
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Competência"
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
            Left            =   10545
            TabIndex        =   105
            Top             =   780
            Width           =   930
         End
         Begin VB.Label Label13 
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
            Left            =   4860
            TabIndex        =   102
            Top             =   180
            Width           =   990
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
            Index           =   1
            Left            =   10530
            TabIndex        =   100
            Top             =   1380
            Width           =   1470
            WordWrap        =   -1  'True
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
            Left            =   3315
            TabIndex        =   99
            Top             =   1380
            Width           =   1050
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
            Left            =   3082
            TabIndex        =   98
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label2 
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
            Left            =   2085
            TabIndex        =   92
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label24 
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
            Left            =   1170
            TabIndex        =   91
            Top             =   780
            Width           =   480
            WordWrap        =   -1  'True
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
            Index           =   0
            Left            =   253
            TabIndex        =   90
            Top             =   780
            Width           =   585
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Height          =   195
            Left            =   8295
            TabIndex        =   72
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ped. compra"
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
            Left            =   10338
            TabIndex        =   71
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label15 
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
            Height          =   195
            Left            =   12030
            TabIndex        =   70
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   0
            Left            =   13658
            TabIndex        =   69
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label21 
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
            Index           =   0
            Left            =   6510
            TabIndex        =   68
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
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
            Left            =   6803
            TabIndex        =   67
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   13035
            TabIndex        =   66
            Top             =   780
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmContas_Pagas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrSql_Contas_Pagas As String 'OK
Public StrSql_Contas_PagasTotal As String 'OK
Dim TBLISTA_Contas_Pagas As ADODB.Recordset 'OK
Public Filtro_Contas_Pagas_Func As String 'OK
Public Filtro_Contas_Pagas_PC As Boolean 'OK
Public Filtro_Contas_Pagas_FuncRel As String 'OK
Public FormulaRel_Contas_Pagas As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=11-CmvxqK4E&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=43&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
lst_ContasPagas.ListItems.Clear
Set TBLISTA_Contas_Pagas = CreateObject("adodb.recordset")
TBLISTA_Contas_Pagas.Open StrSql_Contas_Pagas, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Contas_Pagas.EOF = False Then ProcExibePagina (Pagina)
ProcCarregaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

TotalGeral = 0
Codproduto = 0
TotContas = 0
lst_ContasPagas.ListItems.Clear
TBLISTA_Contas_Pagas.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Contas_Pagas.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Contas_Pagas.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Contas_Pagas.RecordCount - IIf(Pagina > 1, (TBLISTA_Contas_Pagas.PageSize * (Pagina - 1)), 0), TBLISTA_Contas_Pagas.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Contas_Pagas.EOF = False And (ContadorReg <= TamanhoPagina)
    With lst_ContasPagas.ListItems
        .Add , , TBLISTA_Contas_Pagas!IDintconta
        .Item(.Count).SubItems(1) = Format(TBLISTA_Contas_Pagas!Dt_emissao, "dd/mm/yy")
        .Item(.Count).SubItems(2) = Format(TBLISTA_Contas_Pagas!dt_Pagamento, "dd/mm/yy")
        
        If TBLISTA_Contas_Pagas!Parcial = True And TBLISTA_Contas_Pagas!status <> "TÍTULO PAGO PARCIAL LIQUIDADO" Then
            valor = IIf(IsNull(TBLISTA_Contas_Pagas!pagoparcial), 0, TBLISTA_Contas_Pagas!pagoparcial) + IIf(IsNull(TBLISTA_Contas_Pagas!ValorPendente), 0, TBLISTA_Contas_Pagas!ValorPendente)
        Else
            valor = IIf(IsNull(TBLISTA_Contas_Pagas!dbl_valorpagto), 0, TBLISTA_Contas_Pagas!dbl_valorpagto)
        End If
        .Item(.Count).SubItems(3) = Format(valor, "###,##0.00")
        
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Contas_Pagas!txt_ndocumento), "", TBLISTA_Contas_Pagas!txt_ndocumento)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Contas_Pagas!txt_Parcela), "", TBLISTA_Contas_Pagas!txt_Parcela)
        .Item(.Count).SubItems(6) = Trim(TBLISTA_Contas_Pagas!Txt_fornecedor)
        .Item(.Count).SubItems(7) = Format(IIf(IsNull(TBLISTA_Contas_Pagas!ValorPago), 0, TBLISTA_Contas_Pagas!ValorPago), "###,##0.00")
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Contas_Pagas!DataBaixa), "", Format(TBLISTA_Contas_Pagas!DataBaixa, "dd/mm/yy"))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Contas_Pagas!resppag), "", TBLISTA_Contas_Pagas!resppag)
    End With
    TBLISTA_Contas_Pagas.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Contas_Pagas.RecordCount
If TBLISTA_Contas_Pagas.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Contas_Pagas.PageCount
ElseIf TBLISTA_Contas_Pagas.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Pagas.PageCount & " de: " & TBLISTA_Contas_Pagas.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Pagas.AbsolutePage - 1 & " de: " & TBLISTA_Contas_Pagas.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

TotalGeral = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open StrSql_Contas_PagasTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    TotalGeral = IIf(IsNull(TBTotaisnota!TotalGeral), 0, TBTotaisnota!TotalGeral)
End If
TBTotaisnota.Close
TotalContas.Text = Format(TotalGeral, "###,##0.00")

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

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With lst_ContasPagas
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

txtIDFornec = ""
txtFornec = ""
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

Private Sub Cmd_competencia_Click()
On Error GoTo tratar_erro

If txt_Competencia <> "" Then
    ProcFiltrarContas "Competencia = '" & txt_Competencia.Text & "'", "{tbl_ContasPagar.Competencia} = '" & txt_Competencia.Text & "'", True, False, False, False, False, Date, Date, "DataBaixa"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_data_transacao_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "Data_transacao = '" & Format(Txt_data_transacao.Value, "Short Date") & "'", "{tbl_ContasPagar.Data_transacao} = Date(" & Year(Txt_data_transacao.Value) & "," & Month(Txt_data_transacao.Value) & "," & Day(Txt_data_transacao.Value) & ")", True, True, False, False, False, Txt_data_transacao, Txt_data_transacao, "Data_transacao"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcFiltrarContas(TextoFiltro As String, TextoFiltroRel As String, Imprimir As Boolean, DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataBaixa As Boolean, DataInicio As Date, DataFinal As Date, Ordenar As String)
On Error GoTo tratar_erro

ProcConstruirFiltroPadrao TextoFiltro, TextoFiltroRel, Ordenar
ProcSalvarDadosRel DataTransacao, DataEmissao, DataVencimento, DataBaixa, DataInicio, DataFinal
ProcAtualizalista (1)
Imprimir = Imprimir
Filtro_Contas_Pagas_PC = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcConstruirFiltroPadrao(TextoFiltro As String, TextoFiltroRel As String, Ordenar As String)
On Error GoTo tratar_erro

CamposFiltro = "CP.IDintconta, CP.Dt_emissao, CP.dt_Pagamento, CP.Data_transacao, CP.Parcial, CP.Status, CP.pagoparcial, CP.ValorPendente, CP.dbl_valorpagto, CP.txt_ndocumento, CP.txt_Parcela, CP.Txt_fornecedor, CP.ValorPago, CP.DataBaixa, CP.resppag"
If Left(TextoFiltro, 2) = "PN" Then INNERJOINPADRAO = " from tbl_ContasPagar CP INNER JOIN tbl_proposta_nota PN ON PN.ID_nota = CP.ID_nota" Else INNERJOINPADRAO = " from tbl_ContasPagar CP"
INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
INNERJOINTEXTOSUM = "Select Sum(CP.ValorPago) as TotalGeral " & INNERJOINPADRAO
TextoFiltroPadrao = "CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and CP.LogSit = 'S' And " & Filtro_Contas_Pagas_Func
TextoFiltroPadraoRel = "{tbl_contaspagar.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_ContasPagar.LogSit} = 'S' and " & Filtro_Contas_Pagas_FuncRel
OrdenarTexto = " group by " & CamposFiltro & " order by CP." & Ordenar & " desc, CP.IdIntConta"
StrSql_Contas_Pagas = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
StrSql_Contas_PagasTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadrao
FormulaRel_Contas_Pagas = TextoFiltroRel & " and " & TextoFiltroPadraoRel

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

Private Sub Cmd_localizar_contatos_Click()
On Error GoTo tratar_erro

If txtFornec <> "" Then
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = True
    Financeiro_Contas_Receber = False
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

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_valor_Click()
On Error GoTo tratar_erro
    
If txt_ValorDocto.Text <> "" Then
    valor = txt_ValorDocto
    NovoValor = Replace(valor, ",", ".")
    ProcFiltrarContas "dbl_valorpagto = " & NovoValor, "{tbl_ContasPagar.dbl_valorpagto} = " & NovoValor, True, False, False, False, False, Date, Date, "DataBaixa"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_valor_pago_Click()
On Error GoTo tratar_erro
    
If txt_ValorPago <> "" Then
    valor = txt_ValorPago
    NovoValor = Replace(valor, ",", ".")
    ProcFiltrarContas "ValorPago = " & NovoValor, "{tbl_ContasPagar.ValorPago} = " & NovoValor, True, False, False, False, False, Date, Date, "DataBaixa"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdbaixa_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "DataBaixa = '" & Format(txtBaixado.Value, "Short Date") & "'", "{tbl_ContasPagar.DataBaixa} = Date(" & Year(txtBaixado.Value) & "," & Month(txtBaixado.Value) & "," & Day(txtBaixado.Value) & ")", True, False, False, False, True, txtBaixado, txtBaixado, "DataBaixa"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdbanco_Click()
On Error GoTo tratar_erro

If txtBanco.Text <> "" Then
    ProcFiltrarContas "banco = '" & txtBanco.Text & "'", "{tbl_ContasPagar.banco} = '" & txtBanco.Text & "'", True, False, False, False, True, txtBaixado, txtBaixado, "DataBaixa"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarTodas()
On Error GoTo tratar_erro

ProcConstruirFiltroPadrao "CP.IDintconta IS NOT NULL", "Not(IsNull({tbl_ContasPagar.IDintconta}))", "DataBaixa"
ProcSalvarDadosRel False, False, False, Date, Date, Date
ProcAtualizalista (1)
Imprimir = True
Filtro_Contas_Pagas_PC = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmddoc_Click()
On Error GoTo tratar_erro

If txtNDocumento.Text <> "" Then
    ProcFiltrarContas "txt_NDocumento='" & txtNDocumento.Text & "'", "{tbl_ContasPagar.txt_NDocumento} = '" & txtNDocumento.Text & "'", True, False, False, False, False, Date, Date, "DataBaixa"
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

ProcFiltrarContas "dt_emissao = '" & Format(txtDTEmissao.Value, "Short Date") & "'", "{tbl_ContasPagar.dt_emissao} = Date(" & Year(txtDTEmissao.Value) & "," & Month(txtDTEmissao.Value) & "," & Day(txtDTEmissao.Value) & ")", True, False, True, False, False, txtDTEmissao, txtDTEmissao, "dt_Emissao"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPlanoContas()
On Error GoTo tratar_erro
    
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Pagas = True
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

Private Sub ProcCC()
On Error GoTo tratar_erro

If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de cadastrar/visualizar o(s) centro(s) de custo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = True
Faturamento = False
Permitido = True
TextoFiltro = ""
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select ID_nota, txt_ndocumento, Txt_pedido from tbl_contaspagar where IdIntConta = " & IIf(txtidintconta = "", 0, txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    INNERJOINTEXTO = "NF.TipoNF from ((((tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID) INNER JOIN Compras_pedido CP ON CP.Pedido = PN.Proposta) INNER JOIN Compras_pedido_lista CPL ON CPL.IDpedido = CP.IDpedido) LEFT JOIN Compras_pedido_lista_custo CPLC ON CP.IDPedido = CPLC.IDpedido) LEFT JOIN Projproduto P ON P.Desenho = CPL.Desenho"
    If IsNull(TBContas!ID_nota) = False And TBContas!ID_nota <> "" And TBContas!ID_nota <> "0" Then
        TextoFiltro = "PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 2"
    ElseIf IsNull(TBContas!txt_ndocumento) = False And TBContas!txt_ndocumento <> "" Then
            TextoFiltro = "PN.NF = '" & TBContas!txt_ndocumento & "' and NF.int_TipoNota = 2"
        ElseIf IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" And TBContas!Txt_pedido <> "0" Then
                INNERJOINTEXTO = "CP.IDpedido from ((Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CPL.IDpedido = CP.IDpedido) LEFT JOIN Compras_pedido_lista_custo CPLC ON CP.IDPedido = CPLC.IDpedido) INNER JOIN Projproduto P ON P.Desenho = CPL.Desenho"
                TextoFiltro = "CP.Pedido = '" & TBContas!Txt_pedido & "'"
    End If
    If TextoFiltro <> "" Then
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select " & INNERJOINTEXTO & " where " & TextoFiltro & " and (CPLC.ID IS NOT NULL or P.Estoque = 'True')", Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Permitido = False
        End If
    End If
End If
If Permitido = True Then frmContas_CC.Show 1 Else frmContas_pagar_lista_CC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVisualizar()
On Error GoTo tratar_erro

If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de visualizar a lista de antecipações/devoluções relacionadas."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Financeiro_Contas_Pagas = True
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
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

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_fornecedor_Click()
On Error GoTo tratar_erro
    
If txtFornec.Text <> "" Then
    ProcFiltrarContas "txt_fornecedor = '" & txtFornec.Text & "'", "{tbl_ContasPagar.txt_fornecedor} = '" & txtFornec.Text & "'", True, False, False, False, False, Date, Date, "DataBaixa"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmdlocalizarforn_Click()
On Error GoTo tratar_erro

ProcLocalizarFornecedor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarFornecedor()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False
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
If TBLISTA_Contas_Pagas.AbsolutePage <> 2 Then
    If TBLISTA_Contas_Pagas.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Contas_Pagas.PageCount - 1)
    Else
        TBLISTA_Contas_Pagas.AbsolutePage = TBLISTA_Contas_Pagas.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Contas_Pagas.AbsolutePage)
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
    TBLISTA_Contas_Pagas.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Contas_Pagas.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Pagas.AbsolutePage = 1
ProcExibePagina (TBLISTA_Contas_Pagas.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Pagas.AbsolutePage <> -3 Then
    If TBLISTA_Contas_Pagas.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Contas_Pagas.AbsolutePage)
    End If
Else
   ProcExibePagina (TBLISTA_Contas_Pagas.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Pagas.AbsolutePage = TBLISTA_Contas_Pagas.PageCount
ProcExibePagina (TBLISTA_Contas_Pagas.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

Proposta = True
If txtNPedido.Text <> "" Then
    NomeRel = "Contas_pagas.rpt"
    ProcConstruirFiltroPadrao "PN.Proposta = '" & txtNPedido & "'", "{tbl_proposta_nota.proposta} = '" & txtNPedido & "'", "DataBaixa"
    ProcSalvarDadosRel False, False, False, Date, Date, Date
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open StrSql_Contas_Pagas, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        ProcConstruirFiltroPadrao "CP.txt_pedido = '" & txtNPedido & "'", "{tbl_ContasPagar.txt_pedido} = '" & txtNPedido & "'", "DataBaixa"
    End If
    TBAbrir.Close
    Imprimir = True
Else
    ProcFiltrarTodas
End If
ProcAtualizalista (1)
Proposta = False
Filtro_Contas_Pagas_PC = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro
    
If txtStatus.Text <> "" Then
    ProcFiltrarContas "status = '" & txtStatus.Text & "'", "{tbl_ContasPagar.status} = '" & txtStatus.Text & "'", True, False, False, False, False, Date, Date, "DataBaixa"
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
    ProcFiltrarContas "class_conta = '" & cmbtipo_conta.Text & "'", "{tbl_ContasPagar.class_conta} = '" & cmbtipo_conta.Text & "'", True, False, False, False, False, Date, Date, "DataBaixa"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaTipoDocumento()
On Error GoTo tratar_erro

ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'P'"
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
TBFI.Open "Select * from tbl_ContasPagar where IdIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!Banco) = False And TBFI!Banco <> "" Then txtBanco = TBFI!Banco
    If IsNull(TBFI!FormaBaixa) = False And TBFI!FormaBaixa <> "" Then txtFormaPagto = TBFI!FormaBaixa
    If IsNull(TBFI!Class_conta) = False And TBFI!Class_conta <> "" Then cmbtipo_conta = TBFI!Class_conta
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
                
txtidintconta = ""
txtNDocumento = ""
txtNPedido.Clear
txtDTEmissao.Value = Date
txtDataPagto.Value = Date
txtIDFornec = ""
txtFornec = ""
txt_Competencia = ""
txtFormaPagto.ListIndex = -1
txtBaixado.Value = Date
Cmb_data_movimentacao.Value = Date
txtBanco.ListIndex = -1
txtConta = ""
txtparcela = ""
txt_ValorDocto = ""
txt_ValorPago = ""
txt_Ndocto = ""
Txt_dias_atraso = ""
txtjuros = ""
Txt_total_juros = ""
Txt_multa = ""
txtDesconto = ""
txtobs_pgto = ""
cmbtipo_conta.ListIndex = -1
txtStatus.Text = ""
chkConta_fixa.Value = 0
chbparcial.Value = 0
txt_tituloref.Text = ""
CodigoLista = 0
txt_Caminho = ""
Lista_PC.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdvencimento_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "dt_pagamento = '" & Format(txtDataPagto.Value, "Short Date") & "'", "{tbl_ContasPagar.dt_pagamento} = Date(" & Year(txtDataPagto.Value) & "," & Month(txtDataPagto.Value) & "," & Day(txtDataPagto.Value) & ")", True, False, False, True, False, txtDataPagto, txtDataPagto, "dt_Pagamento"

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
    Case vbKeyF7: ProcCC
    Case vbKeyF8: ProcVisualizar
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

Formulario = "Financeiro/Contas pagas"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
Imprimir = False
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
Txt_data_transacao.Value = Date
txtDTEmissao.Value = Date
txtDataPagto.Value = Date
txtBaixado.Value = Date
Cmb_data_movimentacao.Value = Date
Cmb_tipo = "Fornecedor"
Cmb_opcao_lista = "Excluir"
ProcVerifAcessosContasFunc
ProcCarregaComboAno cmbAno, "2005", 1
TabFiltro.Tabs(Month(Date)).Selected = True

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifAcessosContasFunc()
On Error GoTo tratar_erro

'Verifica se o usuário pode visualizar as contas dos funcionários
If FunVerifAcessoContasFunc("Financeiro/Contas pagas/Visualizar contas dos funcionários") = True Then
    Filtro_Contas_Pagas_Func = "CP.txt_fornecedor <> 'Null'"
    Filtro_Contas_Pagas_FuncRel = "{tbl_contaspagar.txt_fornecedor} <> 'Null'"
Else
    Filtro_Contas_Pagas_Func = "CP.Tipo <> 'FU'"
    Filtro_Contas_Pagas_FuncRel = "{tbl_contaspagar.Tipo} <> 'FU'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboBanco()
On Error GoTo tratar_erro

ProcCarregaComboBancoFinanceiro txtBanco, "txt_Descricao <> 'Null' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

ProcCarregaComboFormaPgtoRcbto txtFormaPagto, "Tipo = 'P'"
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Contas pagas"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
ProcVerifAcessosContasFunc

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
With lst_ContasPagas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente cancelar a baixa dessa(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Parcial = True Then
                    ProcSomaRecompra .ListItems(InitFor), TBContas!ValorPago
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
                    TBCorretiva.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo)
                        TBCorretiva.Delete
                    End If
                    TBCorretiva.Close
                    
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contaspagar where IdIntConta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        ValorParcial = TBContas!ValorPago
                        Pendente = TBCorretiva!dbl_valorpagto
                        TBCorretiva!dbl_valorpagto = (Pendente + ValorParcial)
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & IIf(TBContas!tituloref = "", 0, TBContas!tituloref) & "' and IdIntConta <> " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                        Else
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                            TBCorretiva!Parcial = False
                            TBCorretiva!pagoparcial = 0
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
                            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where " & TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                TBFluxo!valor = Format(TBFluxo!valor - Valor3, "###,##0.00")
                                TBFluxo.Update
                                If TBFluxo!valor <= 0 Then TBFluxo.Delete
                            End If
                        End If
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                        TBFluxo!Operacao = "À Debitar"
                        TBFluxo!Data = TBCorretiva!dt_Pagamento
                        TBFluxo!valor = TBCorretiva!dbl_valorpagto
                        TBFluxo!Descricao = TBCorretiva!Txt_fornecedor
                        TBFluxo!status = "N"
                        TBFluxo!int_NotaFiscal = TBContas!txt_ndocumento
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!Cheque = 0
                        TBFluxo!Bloqueado = False
                        TBFluxo.Update
                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                        TBFluxo.Close
                                    
                        If TBContas!FormaBaixa = "SAQUE" Then
                            'Verifica saque e atualiza saldo
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select IDSaque from tbl_ContasPagar_Saque where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                IDlista = TBProduto!IDSaque
                                TBProduto.Delete
                                
                                ProcAtualizaSaldoSaque IDlista
                            End If
                            TBProduto.Close
                        ElseIf TBContas!FormaBaixa <> "CHEQUE" And TBContas!FormaBaixa <> "CHEQUE PRÉ-DATADO" Then
                                'Verifica saldo da antecipação
                                Qtd = .ListItems(InitFor).SubItems(7)
                                Qtde = 0
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "SELECT Sum(Valor) as valor from tbl_Contas_antecipacao where id_conta = " & .ListItems(InitFor) & " and tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = False Then Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                                TBAbrir.Close
                                qt = (Qtd - Qtde) - Qtd_Prog
                                
                                Set TBProduto = CreateObject("adodb.recordset")
                                TBProduto.Open "Select * from tbl_instituicoes where txt_descricao = '" & TBContas!Banco & "' and ID_empresa = " & TBContas!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                If TBProduto.EOF = False Then
                                    TBProduto!Saldo = TBProduto!Saldo + qt
                                    TBProduto.Update
                                End If
                                TBProduto.Close
                        End If
                    End If
                    
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from familia_financeiro where idconta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref) & " and tipoconta = 'P' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        Do While TBFamilia.EOF = False
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from familia_financeiro where IDConta = " & TBContas!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                TBFamilia!valor = TBFamilia!valor + TBCiclo!valor
                                TBFamilia.Update
                            End If
                            TBCiclo.Close
                            
                            Conexao.Execute "DELETE from familia_financeiro where IDConta = " & TBContas!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'P' and Deposito_transf = 'False'"
                            TBFamilia.MoveNext
                        Loop
                    End If
                    
                    'Centro de custo
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from CC_realizado where ID_financeiro = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref) & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        Do While TBFamilia.EOF = False
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from CC_realizado where ID_financeiro = " & .ListItems(InitFor) & " and ID_CC = " & TBFamilia!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                TBFamilia!valor = TBFamilia!valor + TBCiclo!valor
                                TBFamilia!Percentual = TBFamilia!Percentual + TBCiclo!valor
                                TBFamilia.Update
                                TBCiclo.Delete
                            End If
                            TBCiclo.Close
                            TBFamilia.MoveNext
                        Loop
                    End If
                    TBFamilia.Close
                    
                    TBCorretiva.Update
                    TBCorretiva.Close
                Else
                    ProcSomaRecompra .ListItems(InitFor), TBContas!ValorPago
                    If TBContas!Antecipacao = False Then ProcAtualizaSaldoAntecipacao .ListItems(InitFor)
                    If TBContas!Devolucao = True Then procExcluirDevolucao .ListItems(InitFor)
                    
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contaspagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & .ListItems(InitFor) & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                        Else
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                            TBCorretiva!Parcial = False
                            TBCorretiva!pagoparcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                        
                        'Verifica valor do fluxo (normal ou com antecipação)
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select Valor from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
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
                                TextoFiltro = TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Débito' and (idintconta = 0 or idintconta IS NULL)"
                            Else
                                If IsNull(TBContas!ID_varias) = True Or TBContas!ID_varias = 0 Then TextoFiltro = TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Débito'" Else TextoFiltro = "ID_varias = " & TBContas!ID_varias
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
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where idfluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                        TBFluxo!Operacao = "À Debitar"
                        TBFluxo!Data = TBCorretiva!dt_Pagamento
                        TBFluxo!valor = TBCorretiva!dbl_valorpagto
                        TBFluxo!Descricao = TBCorretiva!Txt_fornecedor
                        TBFluxo!status = "N"
                        TBFluxo!int_NotaFiscal = TBContas!txt_ndocumento
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!Cheque = 0
                        TBFluxo!Bloqueado = False
                        TBFluxo.Update
                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                        TBFluxo.Close
                                   
                        TBCorretiva!Logsit = "N"
                        TBCorretiva!DataBaixa = Null
                        TBCorretiva!Data_movimentacao = Null
                        TBCorretiva!Bom_para = Null
                        TBCorretiva!ValorPago = 0
                        TBCorretiva!NDoctoBaixa = ""
                        TBCorretiva!Obs = ""
                        TBCorretiva!Favorecido = ""
                        TBCorretiva!Obscheque = ""
                        TBCorretiva!Dias_atraso = 0
                        TBCorretiva!Juros = 0
                        TBCorretiva!Juros_valor = 0
                        TBCorretiva!Multa = 0
                        TBCorretiva!Multa_valor = 0
                        TBCorretiva!Desconto = 0
                        TBCorretiva!Desconto_valor = 0
                        
                        If TBContas!FormaBaixa = "SAQUE" Then
                            'Verifica saque e atualiza saldo
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select IDSaque from tbl_ContasPagar_Saque where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                IDlista = TBProduto!IDSaque
                                TBProduto.Delete
                                
                                ProcAtualizaSaldoSaque IDlista
                            End If
                            TBProduto.Close
                        ElseIf TBContas!FormaBaixa <> "CHEQUE" And TBContas!FormaBaixa <> "CHEQUE PRÉ-DATADO" Then
                                'Verifica saldo da antecipação
                                Qtd = .ListItems(InitFor).SubItems(7)
                                Qtde = 0
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "SELECT Sum(Valor) as valor from tbl_Contas_antecipacao where id_conta = " & .ListItems(InitFor) & " and tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = False Then
                                    Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                                End If
                                TBAbrir.Close
                                qt = Qtd - Qtde
                        
                                Set TBProduto = CreateObject("adodb.recordset")
                                TBProduto.Open "Select Saldo from tbl_instituicoes where txt_descricao = '" & TBContas!Banco & "' and ID_empresa = " & TBContas!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                If TBProduto.EOF = False Then
                                    If TBContas!Devolucao = True Then TBProduto!Saldo = Format(TBProduto!Saldo - Qtd_Prog, "###,##0.00") Else TBProduto!Saldo = Format(TBProduto!Saldo + qt, "###,##0.00")
                                    TBProduto.Update
                                End If
                                TBProduto.Close
                        End If
                    End If
                    
                    Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where idconta = " & .ListItems(InitFor) & " and tipoconta = 'P'"
                    
                    TBCorretiva!ID_varias = 0
                    TBCorretiva.Update
                    TBCorretiva.Close
                End If
                Conexao.Execute "DELETE from tbl_contas_antecipacao where ID_Conta = " & .ListItems(InitFor) & " and tipo = 'P'"
                
                '==================================
                Modulo = "Financeiro/Contas pagas"
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
    ProcAtualizalista (1)
    lst_ContasPagas.SetFocus
    ProcCarregaDados
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSomaRecompra(IDConta As Long, ValorPago As Double)
On Error GoTo tratar_erro

'Soma valor de recompra no bordero
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select troca_titulo_valores.IDduplicata, troca_titulo_valores.valor_enviado FROM troca_titulo_valores INNER JOIN tbl_ContasPagar ON troca_titulo_valores.n_conta = tbl_ContasPagar.idcontareceber where tbl_ContasPagar.IdIntConta = " & IDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Vlrtotalrecompra from troca_titulo where id = " & TBFI!IDduplicata, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir!Vlrtotalrecompra = TBAbrir!Vlrtotalrecompra + ValorPago
        TBAbrir.Update
    End If
    TBAbrir.Close
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldoAntecipacao(IDConta As Long)
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select CA.* FROM tbl_ContasPagar CP INNER JOIN tbl_contas_antecipacao CA ON CP.IdIntconta = CA.ID_conta where CP.IdIntConta = " & IDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Do While TBFI.EOF = False
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Saldo_antecipacao, LogSit from tbl_ContasPagar where IDintconta = " & TBFI!ID_antecipacao, Conexao, adOpenKeyset, adLockOptimistic
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

frmContas_pagas_localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmContas_pagas_menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Imprimir = False
StrSql_Contas_Pagas = ""
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
Acao = "salvar"
If txtidintconta.Text = "" Then
    NomeCampo = "a conta"
    ProcVerificaAcao
    Exit Sub
End If
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtFormaPagto = "CHEQUE" Or txtFormaPagto = "CHEQUE PRÉ-DATADO" Then
    Cheque = "Cheque n. " & txt_Ndocto
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtBanco & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False' and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Não é permitido alterar a baixa em cheque desta conta, pois o mesmo já está compensado."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contaspagar where IdIntConta = " & txtidintconta.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    ID_varias = IIf(IsNull(TBContas!ID_varias), 0, TBContas!ID_varias)
    If txtFormaPagto = "CHEQUE" Or txtFormaPagto = "CHEQUE PRÉ-DATADO" Or txtFormaPagto = "DOC" Or txtFormaPagto = "TED" Or txtFormaPagto = "MALOTE" Or ID_varias <> 0 Then
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
            TextoFiltro1 = TextoFiltroData & " Operacao = 'Débito' and Descricao = '" & Descricao & "' and Cheque = '" & txt_Ndocto & "' and Instituicao = '" & txtBanco & "'"
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
    TBContas!txt_observacoes = txtObs
    TBContas!Obs = txtobs_pgto.Text
    TBContas!caminho = txt_Caminho
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select Data from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        TBFluxo!Data = Cmb_data_movimentacao
        TBFluxo.Update
    End If
    
    If ID_varias <> 0 Then Conexao.Execute "UPDATE tbl_contaspagar Set Data_movimentacao = '" & Cmb_data_movimentacao & "' where IDintConta <> " & txtidintconta & " and ID_varias = " & ID_varias
    
    TBContas.Update
End If
TBContas.Close
USMsgBox ("Alteração da data da movimentação e observações efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Financeiro/Contas pagas"
ID_documento = txtidintconta
Evento = "Alterar"
Documento = "Documento: " & txtNDocumento
Documento1 = ""
ProcGravaEvento
'==================================
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And lst_ContasPagas.ListItems.Count <> 0 Then
    lst_ContasPagas.SelectedItem = lst_ContasPagas.ListItems(CodigoLista)
    lst_ContasPagas.SetFocus
End If
1:

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
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
    TBAbrir.Open "Select * from tbl_contaspagar where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
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
    TBAbrir.Open "Select Sum(Valor) as Valor from Familia_financeiro where IDConta = " & txtidintconta & " and TipoConta = 'P' and ID <> " & Lista_PC.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ValorPago, Devolucao from tbl_contaspagar where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qt = TBAbrir!ValorPago
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
    Modulo = "Financeiro/Contas pagas"
    Evento = "Alterar valor da conta contábil"
    ID_documento = Lista_PC.SelectedItem
    Documento = "Documento: " & txtNDocumento
    Documento1 = "Código do plano: " & Lista_PC.SelectedItem.ListSubItems(1) & " - Descrição do plano: " & Lista_PC.SelectedItem.ListSubItems(2)
    ProcGravaEvento
    '===================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_ContasPagas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lst_ContasPagas
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    Set TBContas = CreateObject("adodb.recordset")
                    TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBContas.EOF = False Then
                        'Verifica se a conta parcial já está líquidada
                        If TBContas!Parcial = True Then
                            Set TBCorretiva = CreateObject("adodb.recordset")
                            TBCorretiva.Open "Select * from tbl_contaspagar where IdIntConta = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                            If TBCorretiva.EOF = False Then
                                If TBCorretiva!status = "TÍTULO PAGO PARCIAL LIQUIDADO" And TBContas!status <> "TÍTULO PAGO PARCIAL LIQUIDADO" Then GoTo Proximo
                            End If
                            TBCorretiva.Close
                        End If
                        If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                            If IsNull(TBContas!NDoctoBaixa) = False And TBContas!NDoctoBaixa <> "" Then
                                Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False' and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = False Then GoTo Proximo
                                TBAbrir.Close
                            End If
                        End If
                        If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                        
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_instituicoes_transf where IdIntConta = " & .ListItems(InitFor) & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then GoTo Proximo
                        TBAbrir.Close
                        
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_devolucao where ID_Conta = '" & .ListItems(InitFor) & "' and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
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
    ProcOrdenaListView lst_ContasPagas, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_ContasPagas_DblClick()
On Error GoTo tratar_erro

If lst_ContasPagas.ListItems.Count = 0 Then Exit Sub

TextoPedido = ""
TextoPedidoRel = ""
Contador2 = 0
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select ID_nota, Txt_pedido from tbl_ContasPagar where IdIntConta = " & lst_ContasPagas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Set TBProposta = CreateObject("adodb.recordset")
    TBProposta.Open "Select CP.IDpedido from (tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID) INNER JOIN Compras_pedido CP ON CP.Pedido = PN.Proposta where PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
    If TBProposta.EOF = False Then
        Do While TBProposta.EOF = False
            If TextoPedido = "" Then
                TextoPedido = "CP.IDpedido = " & TBProposta!IDpedido
                TextoPedidoRel = "{compras_pedido.IDpedido} = " & TBProposta!IDpedido
            Else
                TextoPedido = TextoPedido & " or CP.IDpedido = " & TBProposta!IDpedido
                TextoPedidoRel = TextoPedidoRel & " or {compras_pedido.IDpedido} = " & TBProposta!IDpedido
            End If
            Contador2 = Contador2 + 1
            TBProposta.MoveNext
        Loop
    ElseIf TBContas!Txt_pedido <> "" And IsNull(TBContas!Txt_pedido) = False Then
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select IDpedido from Compras_pedido where pedido = '" & TBContas!Txt_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Contador2 = 1
            TextoPedido = "CP.IDpedido = " & TBProposta!IDpedido
            TextoPedidoRel = "{compras_pedido.IDpedido} = " & TBProposta!IDpedido
        End If
    End If
    TBProposta.Close
End If
TBContas.Close

If TextoPedido <> "" Then
    Formulario = "Compras/Pedido"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    
    With frmCompras_Pedido
        .Sql_Pedido_Localizar = "Select CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado from Compras_pedido CP LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao where " & TextoPedido
        .FormulaRel_Pedido = TextoPedidoRel
        .listapedido.ListItems.Clear
        .ProcAtualizalistapedido (1)
        
        If Contador2 = 1 Then
            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido CP where " & TextoPedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                .ProcLimpar
                .ProcLimpaCamposItem True
                .ProcLimpaCamposServ True
                .ProcPuxaDados
            End If
            TBCompras_Pedido.Close
        End If
        .SSTab1.Tab = 0
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_ContasPagas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lst_ContasPagas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Cmb_opcao_lista = "Excluir" Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    'Verifica se a conta parcial já está líquidada
                    If TBContas!Parcial = True Then
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select * from tbl_contaspagar where tituloref = " & IIf(TBContas!tituloref = "", 0, TBContas!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            If TBCorretiva!status = "TÍTULO PAGO PARCIAL LIQUIDADO" And TBContas!status <> "TÍTULO PAGO PARCIAL LIQUIDADO" Then
                                USMsgBox ("Não é permitido cancelar a baixa desta conta, pois ela já está líquidada."), vbExclamation, "CAPRIND v5.0"
                                TBContas.Close
                                TBCorretiva.Close
                                .ListItems.Item(InitFor).Checked = False
                                Exit Sub
                            End If
                        End If
                        TBCorretiva.Close
                    End If
                    
                    If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                        If IsNull(TBContas!NDoctoBaixa) = False And TBContas!NDoctoBaixa <> "" Then
                            Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False' and Operacao = 'Débito'", Conexao, adOpenKeyset, adLockOptimistic
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
                    TBAbrir.Open "Select * from tbl_instituicoes_transf where IdIntConta = " & .ListItems(InitFor) & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido cancelar a baixa desta conta por este módulo, pois a mesma é uma tarifa bancária."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                    TBAbrir.Close
                    
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_contas_devolucao where ID_Conta = '" & .ListItems(InitFor) & "' and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
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

Private Sub lst_ContasPagas_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcCarregaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

ProcLimpaCampos
If lst_ContasPagas.ListItems.Count = 0 Then Exit Sub
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contaspagar WHERE IdIntConta = " & lst_ContasPagas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    If IsNull(TBContas!ID_empresa) = False And TBContas!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBContas!ID_empresa
    txtidintconta.Text = TBContas!IDintconta
    Txt_data_transacao.Value = IIf(IsNull(TBContas!Data_transacao), Date, Format(TBContas!Data_transacao, "dd/mm/yyyy"))
    txtNDocumento.Text = IIf(IsNull(TBContas!txt_ndocumento), "", TBContas!txt_ndocumento)
    txtDTEmissao.Value = Format(TBContas!Dt_emissao, "dd/mm/yyyy")
    txtDataPagto.Value = Format(TBContas!dt_Pagamento, "dd/mm/yyyy")
    
    If TBContas!Tipo = "C" Then
        Cmb_tipo = "Cliente"
        txtIDFornec = TBContas!int_codforn
    ElseIf IsNull(TBContas!Tipo) = True Or TBContas!Tipo = "" Or TBContas!Tipo = "FO" Then
            Cmb_tipo = "Fornecedor"
            txtIDFornec = TBContas!int_codforn
        ElseIf TBContas!Tipo = "FU" Then
                Cmb_tipo = "Funcionário"
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select Codigo from Funcionarios where ID = " & TBContas!int_codforn, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = False Then
                    txtIDFornec = TBFornecedor!CODIGO
                End If
            Else
                Cmb_tipo = "Instituição bancária"
                txtIDFornec = TBContas!int_codforn
    End If
    txtFornec.Text = TBContas!Txt_fornecedor
    txt_Competencia = IIf(IsNull(TBContas!Competencia), "", TBContas!Competencia)
    
    txtNPedido.Text = IIf(IsNull(TBContas!Txt_pedido), "", TBContas!Txt_pedido)
    txtStatus.Text = IIf(IsNull(TBContas!status), "", TBContas!status)
    If TBContas!Conta_fixa = True Then chkConta_fixa.Value = 1 Else chkConta_fixa.Value = 0
    txtObs.Text = IIf(IsNull(TBContas!txt_observacoes), "", TBContas!txt_observacoes)
    If TBContas!tituloref <> "" Then txt_tituloref.Text = TBContas!tituloref
    If TBContas!Parcial = True And txtStatus.Text <> "TÍTULO PAGO PARCIAL LIQUIDADO" Then
        chbparcial.Value = 1
        txt_ValorDocto = Format(IIf(IsNull(TBContas!pagoparcial), 0, TBContas!pagoparcial) + IIf(IsNull(TBContas!ValorPendente), 0, TBContas!ValorPendente), "###,##0.00")
    Else
        chbparcial.Value = 0
        txt_ValorDocto = IIf(IsNull(TBContas!dbl_valorpagto), 0, Format(TBContas!dbl_valorpagto, "###,##0.00"))
    End If
    txtobs_pgto.Text = IIf(IsNull(TBContas!Obs), "", TBContas!Obs)
    
    'Dados de Pagamento
    LblDocumento.Caption = "N° documento baixa"
    Select Case txtFormaPagto
        Case "DOC": LblDocumento.Caption = "N° do DOC"
        Case "TED": LblDocumento.Caption = "N° do TED"
        Case "CHEQUE": LblDocumento.Caption = "N° do cheque"
        Case "CHEQUE PRÉ-DATADO": LblDocumento.Caption = "N° do cheque"
        Case "MALOTE": LblDocumento.Caption = "N° do malote"
    End Select
    
    txtBaixado.Value = Format(TBContas!DataBaixa, "dd/mm/yyyy")
    Cmb_data_movimentacao.Value = Format(TBContas!Data_movimentacao, "dd/mm/yyyy")
    txt_Caminho = IIf(IsNull(TBContas!caminho), "", TBContas!caminho)
2:
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select txt_Conta from tbl_Instituicoes where txt_Descricao = '" & TBContas!Banco & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtConta.Text = IIf(IsNull(TBAbrir!txt_Conta), "", TBAbrir!txt_Conta)
        End If
        TBAbrir.Close
        
        If IsNull(TBContas!txt_Parcela) = False And TBContas!txt_Parcela <> "" Then txtparcela.Text = TBContas!txt_Parcela
        
        txt_ValorPago.Text = Format(TBContas!ValorPago, "###,##0.00")
        Txt_dias_atraso = IIf(IsNull(TBContas!Dias_atraso), "", TBContas!Dias_atraso)
        txtjuros.Text = IIf(IsNull(TBContas!Juros_valor), 0, Format(TBContas!Juros_valor, "###,##0.0000000"))
        Txt_total_juros = Format(IIf(IsNull(TBContas!Juros_valor), 0, TBContas!Juros_valor) * IIf(IsNull(TBContas!Dias_atraso), 0, TBContas!Dias_atraso), "###,##0.0000000")
        Txt_multa = IIf(IsNull(TBContas!Multa_valor), 0, Format(TBContas!Multa_valor, "###,##0.0000000"))
        txtDesconto.Text = IIf(IsNull(TBContas!Desconto_valor), 0, Format(TBContas!Desconto_valor, "###,##0.0000000"))
        txt_Ndocto.Text = IIf(IsNull(TBContas!NDoctoBaixa), "", TBContas!NDoctoBaixa)
        ProcCarregaPedido
        CodigoLista = lst_ContasPagas.SelectedItem.index
        
        NomeCampo = "o tipo do documento"
        If IsNull(TBContas!Class_conta) = False And TBContas!Class_conta <> "" Then cmbtipo_conta.Text = TBContas!Class_conta
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
TBLISTA.Open "Select FF.ID, F.Codigo, F.txt_descricao, FF.Valor from Familia_financeiro FF INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDConta = " & txtidintconta & " and FF.Tipoconta = 'P' and FF.Pago_recebido = 'True' and FF.Deposito_transf = 'False' order by F.Codigo", Conexao, adOpenKeyset, adLockOptimistic
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
        Frame4.Top = Frame2.Top + Frame2.Height + 10
        With lst_ContasPagas
            .Top = Frame4.Top + Frame4.Height + 10
            .Height = Frame1.Top - .Top
            If .Visible = True Then .SetFocus
        End With
    Case 1:
        Frame4.Top = Frame3.Top + Frame3.Height
        With lst_ContasPagas
            .Top = Frame4.Top + Frame4.Height
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

If TabFiltro.SelectedItem.key = "Todas" Then
    ProcFiltrarTodas
Else
    M = FunVerificaMes(TabFiltro.SelectedItem.key)
    If OptDomes.Value = True Then ProcConstruirFiltroPadrao "month(DataBaixa)= '" & M & "' and Year(DataBaixa) = '" & cmbAno & "'", "Month ({tbl_ContasPagar.DataBaixa}) = " & M & " and year ({tbl_ContasPagar.DataBaixa})= " & cmbAno, "DataBaixa"
    If OptAteomes.Value = True Then ProcConstruirFiltroPadrao "month (DataBaixa)<= '" & M & "' and Year(DataBaixa) = '" & cmbAno & "'", "Month ({tbl_ContasPagar.DataBaixa}) <= " & M & " and year ({tbl_ContasPagar.DataBaixa})= " & cmbAno, "DataBaixa"
    ProcSalvarDadosRel False, False, False, False, Date, Date
    ProcAtualizalista (1)
    Imprimir = True
    Filtro_Contas_Pagas_PC = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtBanco_Click()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_instituicoes where txt_descricao = '" & txtBanco & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtConta.Text = IIf(IsNull(TBAbrir!txt_Conta), "", TBAbrir!txt_Conta)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFornec_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: ProcLocalizarFornecedor
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaPedido()
On Error GoTo tratar_erro

With txtNPedido
    .Clear
    
    If IsNull(TBContas!ID_nota) = False And TBContas!ID_nota <> "" And TBContas!ID_nota <> "0" Then
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select PN.Proposta from tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID where PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Do While TBProposta.EOF = False
                If IsNull(TBProposta!Proposta) = False Then .AddItem TBProposta!Proposta
                TBProposta.MoveNext
            Loop
        Else
            If IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" And TBContas!Txt_pedido <> "0" Then
                .AddItem TBContas!Txt_pedido
                .Text = TBContas!Txt_pedido
            End If
        End If
        TBProposta.Close
    Else
        If IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" And TBContas!Txt_pedido <> "0" Then
            .AddItem TBContas!Txt_pedido
            .Text = TBContas!Txt_pedido
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDFornec_Change()
On Error GoTo tratar_erro

txtFornec = ""
If txtIDFornec <> "" Then
    If Cmb_tipo <> "Funcionário" Then
        VerifNumero = txtIDFornec
        ProcVerificaNumero
        If VerifNumero = False Then
            txtIDFornec = ""
            txtIDFornec.SetFocus
            Exit Sub
        End If
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    If Cmb_tipo = "Cliente" Then
        TBAbrir.Open "Select NomeRazao from Clientes where idcliente = " & txtIDFornec & " and Prospecto = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then txtFornec.Text = TBAbrir!NomeRazao
    ElseIf Cmb_tipo = "Fornecedor" Then
            TBAbrir.Open "Select Nome_Razao from compras_fornecedores where idcliente = " & txtIDFornec & " and Prospecto = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then txtFornec.Text = TBAbrir!Nome_Razao
        ElseIf Cmb_tipo = "Funcionário" Then
                TBAbrir.Open "Select Nome from Funcionarios where Codigo = '" & txtIDFornec & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtFornec.Text = TBAbrir!Nome
            Else
                TBAbrir.Open "Select Txt_descricao from tbl_Instituicoes where ID = " & txtIDFornec, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtFornec.Text = TBAbrir!Txt_descricao
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDFornec_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: ProcLocalizarFornecedor
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcImprimir
    Case 5: ProcPlanoContas
    Case 6: ProcCC
    Case 7: ProcVisualizar
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
TBAbrir.Open "Select Sum(valor) as valor from tbl_contas_devolucao where Id_devolucao = " & IDConta & " and tipo = 'P' and logsit = 'S'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Qtd_Prog = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
TBAbrir.Close

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_contas_devolucao where ID_Devolucao = " & IDConta & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
Do While TBFI.EOF = False
    If TBFI!Logsit = "N" Then
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from tbl_ContasPagar where IdIntConta = " & TBFI!ID_conta, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            If TBItem!Parcial = False Then
                If TBItem!Bloqueado = False Then TBItem!status = "TÍTULO EM ABERTO"
                TBItem!Parcial = False
                TBItem!pagoparcial = 0
                TBItem!ValorPendente = 0
                TBItem!tituloref = ""
                TBItem!valorprincipal = 0
                TBItem!Logsit = "N"
                TBItem!DataBaixa = Null
                TBItem!Data_movimentacao = Null
                TBItem!Bom_para = Null
                TBItem!ValorPago = 0
                TBItem!NDoctoBaixa = ""
                TBItem!Obs = ""
                TBItem!Favorecido = ""
                TBItem!Obscheque = ""
                TBItem!Dias_atraso = 0
                TBItem!Juros = 0
                TBItem!Juros_valor = 0
                TBItem!Multa = 0
                TBItem!Multa_valor = 0
                TBItem!Desconto = 0
                TBItem!Desconto_valor = 0
                TBItem!ID_varias = 0
                TBItem.Update
                Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where idconta = " & TBFI!ID_conta & " and tipoconta = 'P'"
            Else
                Conexao.Execute "DELETE from F from tbl_Fluxo_de_caixa F INNER JOIN tbl_ContasPagar CP ON CP.IDFluxo = F.IDFluxo where CP.IdIntConta = " & IIf(IsNull(TBFI!ID_conta), 0, TBFI!ID_conta)
                Conexao.Execute "DELETE from tbl_ContasPagar where IdIntConta = " & IIf(IsNull(TBFI!ID_conta), 0, TBFI!ID_conta)
                
                Set TBCorretiva = CreateObject("adodb.recordset")
                TBCorretiva.Open "Select * from tbl_contaspagar where IdIntConta = " & IIf(TBItem!tituloref = "", 0, TBItem!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                If TBCorretiva.EOF = False Then
                    ValorParcial = TBItem!ValorPago
                    Pendente = TBItem!ValorPendente
                    If TBCorretiva!Logsit = "N" Then
                        TBCorretiva!dbl_valorpagto = (Pendente + ValorParcial)
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & IIf(TBItem!tituloref = "", 0, TBItem!tituloref) & "' and IdIntConta <> " & IIf(TBItem!tituloref = "", 0, TBItem!tituloref), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                        Else
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                            TBCorretiva!Parcial = False
                            TBCorretiva!pagoparcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                        
                        Set TBFamilia = CreateObject("adodb.recordset")
                        TBFamilia.Open "Select * from familia_financeiro where idconta = " & IIf(TBItem!tituloref = "", 0, TBItem!tituloref) & " and tipoconta = 'P' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFamilia.EOF = False Then
                            Do While TBFamilia.EOF = False
                                Set TBCiclo = CreateObject("adodb.recordset")
                                TBCiclo.Open "Select * from familia_financeiro where IDConta = " & TBItem!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBCiclo.EOF = False Then
                                    TBFamilia!valor = TBFamilia!valor + TBCiclo!valor
                                    TBFamilia.Update
                                End If
                                TBCiclo.Close
                                
                                Conexao.Execute "DELETE from familia_financeiro where IDConta = " & TBItem!IDintconta & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'P' and Deposito_transf = 'False'"
                                TBFamilia.MoveNext
                            Loop
                        End If
                    Else
                        TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                        TBCorretiva!Parcial = True
                        TBCorretiva!pagoparcial = TBCorretiva!dbl_valorpagto
                        TBCorretiva!ValorPendente = Format(TBCorretiva!valorprincipal - TBCorretiva!ValorPago, "###,##0.00")
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
Conexao.Execute "DELETE from tbl_contas_devolucao where ID_Devolucao = " & IDConta & " and Tipo = 'P'"

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
