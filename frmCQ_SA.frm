VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCQ_SA 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Solicitação de ação"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WhatsThisHelp   =   -1  'True
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   93
      Top             =   9720
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "Solicitação de ação"
      TabPicture(0)   =   "frmCQ_SA.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Equipe de trabalho"
      TabPicture(1)   =   "frmCQ_SA.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ListView2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtID1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Outros"
      TabPicture(2)   =   "frmCQ_SA.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(1)=   "SSTab2"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   1155
         Left            =   60
         TabIndex        =   123
         Top             =   8520
         Width           =   15195
         Begin VB.CommandButton cmdSalvarRespSA 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   9570
            Picture         =   "frmCQ_SA.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   510
            Width           =   405
         End
         Begin VB.TextBox txtresponsavelSA 
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
            Left            =   5670
            MaxLength       =   50
            TabIndex        =   124
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   510
            Width           =   3870
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável pela SA"
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
            Left            =   6855
            TabIndex        =   125
            Top             =   300
            Width           =   1500
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   110
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
         ButtonCaption2  =   "Excluir"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Excluir (F4)"
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
         ButtonWidth2    =   45
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Relatório"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Relatório (F5)"
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
         ButtonLeft3     =   95
         ButtonTop3      =   2
         ButtonWidth3    =   60
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior."
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
         ButtonLeft4     =   157
         ButtonTop4      =   2
         ButtonWidth4    =   55
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Próximo"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Próximo registro."
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
         ButtonLeft5     =   214
         ButtonTop5      =   2
         ButtonWidth5    =   55
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
         ButtonLeft6     =   271
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   275
         ButtonTop7      =   2
         ButtonWidth7    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   318
         ButtonTop8      =   2
         ButtonWidth8    =   30
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonKey9      =   "9"
         ButtonAlignment9=   2
         BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState9    =   5
         ButtonLeft9     =   350
         ButtonTop9      =   2
         ButtonWidth9    =   24
         ButtonHeight9   =   24
         ButtonUseMaskColor9=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   12750
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCQ_SA.frx":00A7
            Count           =   1
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   94
         Top             =   9090
         Width           =   15195
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
            TabIndex        =   97
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
            Left            =   2880
            TabIndex        =   96
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
            ItemData        =   "frmCQ_SA.frx":4B6A
            Left            =   6840
            List            =   "frmCQ_SA.frx":4B74
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   180
            Width           =   1965
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   98
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SA.frx":4B8C
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
            TabIndex        =   99
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SA.frx":8330
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
            TabIndex        =   100
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
            TabIndex        =   101
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SA.frx":BE39
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
            TabIndex        =   102
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCQ_SA.frx":FF28
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
         Begin VB.Label Label35 
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
            Left            =   3510
            TabIndex        =   111
            Top             =   240
            Width           =   1440
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
            TabIndex        =   106
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
            TabIndex        =   105
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label33 
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
            Left            =   2190
            TabIndex        =   104
            Top             =   240
            Width           =   645
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
            Left            =   5520
            TabIndex        =   103
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.TextBox txtID1 
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
         Left            =   1470
         Locked          =   -1  'True
         MouseIcon       =   "frmCQ_SA.frx":137B4
         MousePointer    =   99  'Custom
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3420
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   49
         Top             =   1200
         Width           =   11820
         Begin VB.TextBox txtIDContato 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Left            =   1800
            MaxLength       =   60
            MouseIcon       =   "frmCQ_SA.frx":13ABE
            MousePointer    =   99  'Custom
            TabIndex        =   53
            ToolTipText     =   "Digite o nome para contato."
            Top             =   240
            Visible         =   0   'False
            Width           =   950
         End
         Begin VB.TextBox txtNomeContato 
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
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmCQ_SA.frx":13DC8
            MousePointer    =   99  'Custom
            TabIndex        =   52
            ToolTipText     =   "Nome do contato."
            Top             =   240
            Width           =   9855
         End
         Begin VB.TextBox txtdepartamento 
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
            Left            =   1770
            MaxLength       =   60
            MouseIcon       =   "frmCQ_SA.frx":140D2
            MousePointer    =   99  'Custom
            TabIndex        =   51
            ToolTipText     =   "Departamento do contato."
            Top             =   630
            Width           =   9855
         End
         Begin VB.TextBox TxtEmail_Contato 
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
            Height          =   330
            Left            =   1770
            MouseIcon       =   "frmCQ_SA.frx":143DC
            MousePointer    =   99  'Custom
            TabIndex        =   50
            ToolTipText     =   "E-mail do cliente."
            Top             =   1440
            Width           =   9855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
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
            Left            =   600
            TabIndex        =   57
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do contato:"
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
            Left            =   405
            TabIndex        =   56
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ramal:"
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
            Index           =   13
            Left            =   1200
            TabIndex        =   55
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
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
            Index           =   16
            Left            =   1215
            TabIndex        =   54
            Top             =   1478
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2145
         Left            =   -74925
         TabIndex        =   58
         Top             =   1305
         Width           =   15225
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
            Left            =   6810
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1755
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
            Left            =   8580
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   3300
         End
         Begin VB.ComboBox cmbStatus 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "frmCQ_SA.frx":146E6
            Left            =   13470
            List            =   "frmCQ_SA.frx":146F0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1575
         End
         Begin VB.ComboBox cmbEficaz 
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
            ItemData        =   "frmCQ_SA.frx":14707
            Left            =   14130
            List            =   "frmCQ_SA.frx":14711
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Eficaz."
            Top             =   990
            Width           =   915
         End
         Begin VB.ComboBox cmbTipo 
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
            ItemData        =   "frmCQ_SA.frx":1471F
            Left            =   11895
            List            =   "frmCQ_SA.frx":14729
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Tipo."
            Top             =   390
            Width           =   1565
         End
         Begin VB.TextBox txtObjetivo 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   255
            TabIndex        =   7
            ToolTipText     =   "Objetivo."
            Top             =   990
            Width           =   11055
         End
         Begin VB.TextBox txtID 
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Número da solicitação de ação."
            Top             =   390
            Width           =   1325
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
            Height          =   315
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1095
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
            Height          =   315
            Left            =   3495
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3300
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
            Height          =   435
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Observação."
            Top             =   1590
            Width           =   14865
         End
         Begin MSComCtl2.DTPicker txtPrev 
            Height          =   315
            Left            =   11250
            TabIndex        =   8
            ToolTipText     =   "Data de previsão."
            Top             =   990
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
            Format          =   197984259
            CurrentDate     =   39057
         End
         Begin MSMask.MaskEdBox txtfim 
            Height          =   315
            Left            =   12660
            TabIndex        =   9
            ToolTipText     =   "Data de fechamento."
            Top             =   990
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox txtRNC 
            Height          =   315
            Left            =   1500
            TabIndex        =   121
            ToolTipText     =   "Número da RNC."
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº RNC*"
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
            Left            =   1665
            TabIndex        =   122
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável pela validação"
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
            Left            =   9240
            TabIndex        =   108
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data/hora validação"
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
            TabIndex        =   107
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   14025
            TabIndex        =   69
            Top             =   180
            Width           =   465
         End
         Begin VB.Image Imgcalendario 
            Height          =   360
            Left            =   13725
            Picture         =   "frmCQ_SA.frx":14744
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   960
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo*"
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
            Left            =   12510
            TabIndex        =   68
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Objetivo*"
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
            Left            =   5400
            TabIndex        =   67
            Top             =   780
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Previsão"
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
            Left            =   11640
            TabIndex        =   66
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim"
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
            Left            =   13072
            TabIndex        =   65
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eficaz"
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
            Left            =   14377
            TabIndex        =   64
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº SA"
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
            Left            =   600
            TabIndex        =   63
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   4688
            TabIndex        =   62
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data"
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
            Left            =   2730
            TabIndex        =   61
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -11580
            TabIndex        =   60
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   7140
            TabIndex        =   59
            Top             =   1380
            Width           =   945
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5615
         Left            =   -74925
         TabIndex        =   12
         Top             =   3460
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   9895
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
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "N° SA"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RNC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   15849
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Validado"
            Object.Width           =   1499
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6360
         Left            =   75
         TabIndex        =   15
         Top             =   2175
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11218
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
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   12793
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Setor"
            Object.Width           =   12793
         EndProperty
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   855
         Left            =   75
         TabIndex        =   72
         Top             =   1305
         Width           =   15195
         Begin VB.TextBox Txt_setor 
            BackColor       =   &H00FFFFFF&
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
            Left            =   9180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Setor."
            Top             =   390
            Width           =   5805
         End
         Begin VB.ComboBox Cmb_responsavel 
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
            ItemData        =   "frmCQ_SA.frx":14BC7
            Left            =   180
            List            =   "frmCQ_SA.frx":14BD1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Responsável."
            Top             =   390
            Width           =   8985
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setor"
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
            Left            =   11887
            TabIndex        =   75
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável*"
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
            Left            =   4170
            TabIndex        =   74
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H80000001&
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -11580
            TabIndex        =   73
            Top             =   4200
            Width           =   270
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8385
         Left            =   -74925
         TabIndex        =   77
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   14790
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
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
         TabCaption(0)   =   "Descrição da N/C"
         TabPicture(0)   =   "frmCQ_SA.frx":14BEC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Causa/raiz da N/C"
         TabPicture(1)   =   "frmCQ_SA.frx":14C08
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Ação corretiva"
         TabPicture(2)   =   "frmCQ_SA.frx":14C24
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame7"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Acompanhamento"
         TabPicture(3)   =   "frmCQ_SA.frx":14C40
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame8"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Riscos e oportunidades"
         TabPicture(4)   =   "frmCQ_SA.frx":14C5C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame4"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Fechamento"
         TabPicture(5)   =   "frmCQ_SA.frx":14C78
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame9"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Revisar documentos"
         TabPicture(6)   =   "frmCQ_SA.frx":14C94
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame10"
         Tab(6).Control(1)=   "optSim"
         Tab(6).Control(2)=   "optNao"
         Tab(6).ControlCount=   3
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   -74970
            TabIndex        =   118
            Top             =   330
            Width           =   15135
            Begin VB.CommandButton cmdResponsavel_riscos 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14610
               Picture         =   "frmCQ_SA.frx":14CB0
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Localizar responsável."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtResponsavel_riscos 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   35
               TabStop         =   0   'False
               ToolTipText     =   "Responsável."
               Top             =   390
               Width           =   12915
            End
            Begin VB.TextBox txtRiscos 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7125
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               ToolTipText     =   "Descrição."
               Top             =   780
               Width           =   14745
            End
            Begin MSMask.MaskEdBox txtData_riscos 
               Height          =   315
               Left            =   180
               TabIndex        =   34
               ToolTipText     =   "Data."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Image cmdData_riscos 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":14DB2
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   120
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   7680
               TabIndex        =   119
               Top             =   180
               Width           =   915
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   -74970
            TabIndex        =   115
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtTexto5 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7125
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   41
               ToolTipText     =   "Descrição."
               Top             =   780
               Width           =   14745
            End
            Begin VB.TextBox txtResponsavel5 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "Responsável."
               Top             =   390
               Width           =   12915
            End
            Begin VB.CommandButton cmdResponsave5 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14610
               Picture         =   "frmCQ_SA.frx":15235
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Localizar responsável."
               Top             =   390
               Width           =   315
            End
            Begin MSMask.MaskEdBox txtData5 
               Height          =   315
               Left            =   180
               TabIndex        =   38
               ToolTipText     =   "Data."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   7680
               TabIndex        =   117
               Top             =   180
               Width           =   915
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   116
               Top             =   180
               Width           =   345
            End
            Begin VB.Image cmdData5 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":15337
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   360
               Width           =   330
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   -74970
            TabIndex        =   87
            Top             =   330
            Width           =   15135
            Begin VB.CommandButton cmdResponsave4 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   13080
               Picture         =   "frmCQ_SA.frx":157BA
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Localizar responsável."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtResponsavel4 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "Responsável."
               Top             =   390
               Width           =   11385
            End
            Begin VB.TextBox txtTexto4 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7125
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   33
               ToolTipText     =   "Descrição."
               Top             =   780
               Width           =   14745
            End
            Begin MSMask.MaskEdBox txtData4 
               Height          =   315
               Left            =   180
               TabIndex        =   29
               ToolTipText     =   "Data."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin MSMask.MaskEdBox txtData7 
               Height          =   315
               Left            =   13530
               TabIndex        =   32
               ToolTipText     =   "Prazo para fechamento da solicitação de ação."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Image cmdData7 
               Height          =   360
               Left            =   14595
               Picture         =   "frmCQ_SA.frx":158BC
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário"
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Prazo SA"
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
               Left            =   13740
               TabIndex        =   91
               Top             =   180
               Width           =   645
            End
            Begin VB.Image cmdData4 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":15D3F
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   89
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   6915
               TabIndex        =   88
               Top             =   180
               Width           =   915
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   -74970
            TabIndex        =   84
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtTexto3 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7125
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   28
               ToolTipText     =   "Descrição."
               Top             =   780
               Width           =   14745
            End
            Begin VB.CommandButton cmdResponsavel3 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   13080
               Picture         =   "frmCQ_SA.frx":161C2
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Localizar responsável."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtResponsavel3 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "Responsável."
               Top             =   390
               Width           =   11385
            End
            Begin MSMask.MaskEdBox txtData3 
               Height          =   315
               Left            =   180
               TabIndex        =   24
               ToolTipText     =   "Data."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin MSMask.MaskEdBox txtData6 
               Height          =   315
               Left            =   13530
               TabIndex        =   27
               ToolTipText     =   "Prazo de implementação."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Image cmdData6 
               Height          =   360
               Left            =   14595
               Picture         =   "frmCQ_SA.frx":162C4
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário"
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pr. de implem."
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
               Left            =   13552
               TabIndex        =   90
               Top             =   180
               Width           =   1020
            End
            Begin VB.Image cmdData3 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":16747
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   86
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   6915
               TabIndex        =   85
               Top             =   180
               Width           =   915
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   -74970
            TabIndex        =   81
            Top             =   330
            Width           =   15135
            Begin VB.CommandButton cmdResponsavel2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14610
               Picture         =   "frmCQ_SA.frx":16BCA
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Localizar responsável."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtResponsavel2 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   21
               TabStop         =   0   'False
               ToolTipText     =   "Responsável."
               Top             =   390
               Width           =   12915
            End
            Begin VB.TextBox txttexto2 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7125
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               ToolTipText     =   "Descrição."
               Top             =   780
               Width           =   14745
            End
            Begin MSMask.MaskEdBox txtData2 
               Height          =   315
               Left            =   180
               TabIndex        =   20
               ToolTipText     =   "Data."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Image cmdData2 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":16CCC
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   83
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   7680
               TabIndex        =   82
               Top             =   180
               Width           =   915
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   30
            TabIndex        =   78
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtResponsavel1 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   17
               TabStop         =   0   'False
               ToolTipText     =   "Responsável."
               Top             =   390
               Width           =   12915
            End
            Begin VB.TextBox txtTexto1 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7125
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               ToolTipText     =   "Descrição."
               Top             =   780
               Width           =   14745
            End
            Begin VB.CommandButton cmdResponsavel1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14610
               Picture         =   "frmCQ_SA.frx":1714F
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Localizar responsável."
               Top             =   390
               Width           =   315
            End
            Begin MSMask.MaskEdBox txtData1 
               Height          =   315
               Left            =   180
               TabIndex        =   16
               ToolTipText     =   "Data."
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Image cmdData1 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":17251
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   80
               Top             =   180
               Width           =   345
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   7680
               TabIndex        =   79
               Top             =   180
               Width           =   915
            End
         End
         Begin VB.OptionButton optNao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não"
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
            Left            =   -74010
            TabIndex        =   43
            Top             =   330
            Width           =   615
         End
         Begin VB.OptionButton optSim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sim"
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
            Left            =   -74730
            TabIndex        =   42
            Top             =   330
            Width           =   555
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   -74970
            TabIndex        =   112
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtTexto6 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7035
               Left            =   180
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               ToolTipText     =   "Descrição."
               Top             =   870
               Width           =   14745
            End
            Begin VB.CommandButton cmdResponsavel 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14610
               Picture         =   "frmCQ_SA.frx":176D4
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Localizar responsável."
               Top             =   450
               Width           =   315
            End
            Begin VB.TextBox txtResponsavel_revisar 
               BackColor       =   &H00FFFFFF&
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
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pela revisão."
               Top             =   450
               Width           =   12915
            End
            Begin MSMask.MaskEdBox txtData_revisar 
               Height          =   315
               Left            =   180
               TabIndex        =   44
               ToolTipText     =   "Data da revisão."
               Top             =   450
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
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
               Left            =   540
               TabIndex        =   114
               Top             =   240
               Width           =   345
            End
            Begin VB.Image cmdData 
               Height          =   360
               Left            =   1245
               Picture         =   "frmCQ_SA.frx":177D6
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   420
               Width           =   330
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   7680
               TabIndex        =   113
               Top             =   240
               Width           =   915
            End
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   92
         Top             =   330
         Width           =   15195
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
         ButtonCaption8  =   "Filtrar todos"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Filtrar todas solicitações."
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
         ButtonWidth8    =   66
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
         ButtonLeft9     =   415
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
         ButtonLeft10    =   470
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
         ButtonLeft11    =   474
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
         ButtonLeft12    =   517
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
         ButtonLeft13    =   549
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12750
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCQ_SA.frx":17C59
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   75
         TabIndex        =   109
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
         ButtonLeft2     =   40
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
         ButtonLeft3     =   86
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
         ButtonLeft4     =   133
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
         ButtonLeft5     =   195
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
         ButtonLeft6     =   252
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
         ButtonLeft7     =   309
         ButtonTop7      =   4
         ButtonWidth7    =   2
         ButtonHeight7   =   54
         ButtonCaption8  =   "Ajuda"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Ajuda (F1)"
         ButtonKey8      =   "9"
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
         ButtonLeft8     =   313
         ButtonTop8      =   2
         ButtonWidth8    =   41
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Sair"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Sair (Esc)"
         ButtonKey9      =   "10"
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
         ButtonLeft9     =   356
         ButtonTop9      =   2
         ButtonWidth9    =   30
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonKey10     =   "11"
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
         ButtonLeft10    =   388
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   12750
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCQ_SA.frx":1EFBA
            Count           =   1
         End
      End
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data :"
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
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Responsável :"
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
      Left            =   1560
      TabIndex        =   70
      Top             =   0
      Width           =   1020
   End
End
Attribute VB_Name = "frmCQ_SA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_SA                   As Boolean 'OK
Dim Novo_SA1                  As Boolean 'OK
Dim CQ_SA_Alteracao           As Boolean
Dim CQ_SA_Aba                 As Integer
Dim CQ_SA_Texto               As String
Public Responsavel            As Integer 'OK
Public StrSql_CQ_SA_Localizar As String 'OK
Public TBLISTA_CQ_SA          As ADODB.Recordset 'OK

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With ListView1
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    Select Case Cmb_opcao_lista
        Case "Validação"
            .ButtonState(4) = 5
            .ButtonState(9) = 0
        Case "Excluir"
            .ButtonState(4) = 0
            .ButtonState(9) = 5
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_responsavel_Click()
On Error GoTo tratar_erro

If Cmb_responsavel = "" Then Exit Sub
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select Setor from Usuarios where Nome = '" & Cmb_responsavel & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    Txt_setor = IIf(IsNull(TBUsuarios!Setor), "", TBUsuarios!Setor)
End If
TBUsuarios.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_SA order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos2
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_SA where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        ProcCarregaLista1
        ProcPuxadados2
    Else
        USMsgBox ("Fim dos cadastros de solicitação de ação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_SA = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrarTodos()
On Error GoTo tratar_erro

StrSql_CQ_SA_Localizar = "Select * from CQ_SA order by id desc"
ProcCarregaLista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData_Click()
On Error GoTo tratar_erro

If optNao.Value = True Then Exit Sub
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
CQ_SA_Texto = txtData_revisar
CQ_SA_Aba = 21
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData_riscos_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
Sit_Data = 10
CQ_SA_Texto = txtData_riscos
CQ_SA_Aba = 15
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData1_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
CQ_SA_Texto = txtData1
CQ_SA_Aba = 1
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData2_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
CQ_SA_Texto = txtData2
CQ_SA_Aba = 4
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData3_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
Sit_Data = 5
CQ_SA_Texto = txtData3
CQ_SA_Aba = 7
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData4_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
Sit_Data = 6
CQ_SA_Texto = txtData4
CQ_SA_Aba = 11
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData5_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
Sit_Data = 7
CQ_SA_Texto = txtData5
CQ_SA_Aba = 18
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData6_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
Sit_Data = 8
CQ_SA_Texto = txtData6
CQ_SA_Aba = 9
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdData7_Click()
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
Funcionario = False
RNC = False
SolicitacaoAcao = True
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
Sit_Data = 9
CQ_SA_Texto = txtdata7
CQ_SA_Aba = 13
FrmCalendario.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_SA order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimpaCampos2
        txtId = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from CQ_SA where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcPuxaDados
        End If
        ProcPuxadados2
        ProcCarregaLista1
    Else
        USMsgBox ("Fim dos cadastros de solicitação de ação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_SA = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CQ_SA.AbsolutePage <> 2 Then
    If TBLISTA_CQ_SA.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_CQ_SA.PageCount - 1)
    Else
        TBLISTA_CQ_SA.AbsolutePage = TBLISTA_CQ_SA.AbsolutePage - 2
        ProcExibePagina (TBLISTA_CQ_SA.AbsolutePage)
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
    TBLISTA_CQ_SA.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_CQ_SA.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CQ_SA.AbsolutePage = 1
ProcExibePagina (TBLISTA_CQ_SA.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CQ_SA.AbsolutePage <> -3 Then
    If TBLISTA_CQ_SA.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_CQ_SA.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_CQ_SA.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CQ_SA.AbsolutePage = TBLISTA_CQ_SA.PageCount
ProcExibePagina (TBLISTA_CQ_SA.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsave4_Click()
On Error GoTo tratar_erro
  
Responsavel = 4
CQ_SA_Texto = txtResponsavel4
CQ_SA_Aba = 12
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsave5_Click()
On Error GoTo tratar_erro
  
Responsavel = 5
CQ_SA_Texto = txtResponsavel5
CQ_SA_Aba = 19
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsavel_Click()
On Error GoTo tratar_erro
  
If optNao.Value = True Then Exit Sub
Responsavel = 0
CQ_SA_Texto = txtResponsavel_revisar
CQ_SA_Aba = 22
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsavel_riscos_Click()
On Error GoTo tratar_erro
  
Responsavel = 6
CQ_SA_Texto = txtResponsavel_riscos
CQ_SA_Aba = 16
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsavel1_Click()
On Error GoTo tratar_erro
  
Responsavel = 1
CQ_SA_Texto = txtResponsavel1
CQ_SA_Aba = 2
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsavel2_Click()
On Error GoTo tratar_erro
  
Responsavel = 2
CQ_SA_Texto = txtResponsavel2
CQ_SA_Aba = 5
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdResponsavel3_Click()
On Error GoTo tratar_erro
  
Responsavel = 3
CQ_SA_Texto = txtResponsavel3
CQ_SA_Aba = 8
frmCQ_SA_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdSalvarRespSA_Click()
On Error GoTo tratar_erro
  
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * from CQ_SA where id = '" & txtId & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    TBUsuarios!ResponsavelSA = txtresponsavelSA
    TBUsuarios.Update
End If
TBUsuarios.Close

MsgBox ("O responsável da SA foi incluído com sucesso!"), vbInformation + vbOKOnly


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
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros ListView1, "Qualidade/Solicitação de ação"
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo1
            Case vbKeyF3: procSalvar1
            Case vbKeyF4: ProcExcluir1
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF3: procSalvar2
            Case vbKeyF4: procExcluir2
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Caption = "Qualidade - Solicitação de ação (Nº SA : " & TBAbrir!ID & ")"
txtId = TBAbrir!ID
txtRNC = IIf(IsNull(TBAbrir!RNC), "____-__", (TBAbrir!RNC))
txtData = IIf(IsNull(TBAbrir!Data), "", (Format(TBAbrir!Data, "dd/mm/yy")))
txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", (TBAbrir!Responsavel))
If IsNull(TBAbrir!Tipo) = False And TBAbrir!Tipo <> "" Then cmbTipo = TBAbrir!Tipo
txtObjetivo = IIf(IsNull(TBAbrir!Objetivo), "", TBAbrir!Objetivo)
txtSetor = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
If IsNull(TBAbrir!eficaz) = False And TBAbrir!eficaz <> "" Then cmbEficaz = TBAbrir!eficaz
txtfim = IIf(IsNull(TBAbrir!FIM), "__/__/____", Format(TBAbrir!FIM, "dd/mm/yyyy"))
txtPrev = IIf(IsNull(TBAbrir!Previsao), Date, TBAbrir!Previsao)
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
If IsNull(TBAbrir!status) = False And TBAbrir!status <> "" Then cmbStatus = TBAbrir!status
If TBAbrir!Revisar = True Then
    optSim.Value = True
    txtData_revisar = IIf(IsNull(TBAbrir!Data_revisar), "__/__/____", Format(TBAbrir!Data_revisar, "DD/mm/yyyy"))
    txtResponsavel_revisar = IIf(IsNull(TBAbrir!Responsavel_revisar), "", TBAbrir!Responsavel_revisar)
Else
    optNao.Value = True
    txtData_revisar = "__/__/____"
    txtResponsavel_revisar = ""
End If
txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
txtresponsavelSA = IIf(IsNull(TBAbrir!ResponsavelSA), "", TBAbrir!ResponsavelSA)
Novo_SA = False
Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = ""
txtData.Text = Format(Date, "dd/mm/yy")
txtResponsavel.Text = pubUsuario
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select Setor from Usuarios where Usuario = '" & txtResponsavel & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    txtSetor = TBUsuarios!Setor
End If
TBUsuarios.Close
cmbTipo.ListIndex = -1
cmbStatus.ListIndex = -1
txtObjetivo = ""
txtPrev = Date
txtfim = "__/__/____"
cmbEficaz.ListIndex = -1
txtObs = ""
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtRNC = "____-__"
ProcLimpaCampos2
CodigoLista = 0
Caption = "Qualidade - Solicitação de ação"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos1()
On Error GoTo tratar_erro

txtId1 = 0
Cmb_responsavel.ListIndex = -1
Txt_setor = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

txtTexto1 = ""
txtTexto2 = ""
txtTexto3 = ""
txtTexto4 = ""
txtTexto5 = ""
txtTexto6 = ""
txtRiscos = ""
txtData1 = "__/__/____"
txtData2 = "__/__/____"
txtData3 = "__/__/____"
txtData4 = "__/__/____"
txtData5 = "__/__/____"
txtData_revisar = "__/__/____"
txtData_riscos = "__/__/____"
txtResponsavel1 = ""
txtResponsavel2 = ""
txtResponsavel3 = ""
txtResponsavel4 = ""
txtResponsavel5 = ""
txtResponsavel_revisar = ""
txtResponsavel_riscos = ""
optSim.Value = False
optNao.Value = False
CQ_SA_Alteracao = False
txtresponsavelSA.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados2()
On Error GoTo tratar_erro

TBGravar!Texto1 = Trim(txtTexto1)
TBGravar!Data1 = IIf(txtData1 = "__/__/____", Null, txtData1)
TBGravar!Responsavel1 = txtResponsavel1
TBGravar!Texto2 = Trim(txtTexto2)
TBGravar!Data2 = IIf(txtData2 = "__/__/____", Null, txtData2)
TBGravar!responsavel2 = txtResponsavel2
TBGravar!Texto3 = Trim(txtTexto3)
TBGravar!Data3 = IIf(txtData3 = "__/__/____", Null, txtData3)
TBGravar!Data6 = IIf(txtData6 = "__/__/____", Null, txtData6)
TBGravar!responsavel3 = txtResponsavel3
TBGravar!texto4 = Trim(txtTexto4)
TBGravar!Data4 = IIf(txtData4 = "__/__/____", Null, txtData4)
TBGravar!Data7 = IIf(txtdata7 = "__/__/____", Null, txtdata7)
TBGravar!responsavel4 = txtResponsavel4
TBGravar!Texto5 = Trim(txtTexto5)
TBGravar!Data5 = IIf(txtData5 = "__/__/____", Null, txtData5)
TBGravar!responsavel5 = txtResponsavel5
TBGravar!Riscos = Trim(txtRiscos)
TBGravar!Data_riscos = IIf(txtData_riscos = "__/__/____", Null, txtData_riscos)
TBGravar!Responsavel_riscos = txtResponsavel_riscos
If optSim.Value = True Then
    TBGravar!Revisar = True
    TBGravar!Data_revisar = IIf(txtData_revisar = "__/__/____", Null, txtData_revisar)
    TBGravar!Responsavel_revisar = IIf(txtResponsavel_revisar = "", Null, txtResponsavel_revisar)
    TBGravar!Texto6 = Trim(txtTexto6)
Else
    TBGravar!Revisar = False
    TBGravar!Data_revisar = Null
    TBGravar!Responsavel_revisar = ""
    TBGravar!Texto6 = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar2()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "solicitação de ação", "os dados", False) = False Then Exit Sub

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_SA WHERE ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviaDados2
TBGravar.Update
TBGravar.Close
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Qualidade/RNC"
Evento = "Alterar outros"
ID_documento = txtId
Documento = "Nº SA: " & txtId
Documento1 = ""
ProcGravaEvento
'==================================
CQ_SA_Alteracao = False
CQ_SA_Aba = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir2()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If FunVerifValidacaoRegistro("excluir", txtDtValidacao, "solicitação de ação", "os dados", False) = False Then Exit Sub

If USMsgBox("Deseja realmente excluir?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from CQ_SA where id  = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        TBGravar!Texto1 = Null
        TBGravar!Texto2 = Null
        TBGravar!Texto3 = Null
        TBGravar!texto4 = Null
        TBGravar!Texto5 = Null
        TBGravar!Riscos = Null
        TBGravar!Data1 = Null
        TBGravar!Data2 = Null
        TBGravar!Data3 = Null
        TBGravar!Data4 = Null
        TBGravar!Data5 = Null
        TBGravar!Data_riscos = Null
        TBGravar!Responsavel1 = Null
        TBGravar!responsavel2 = Null
        TBGravar!responsavel3 = Null
        TBGravar!responsavel4 = Null
        TBGravar!responsavel5 = Null
        TBGravar!Responsavel_riscos = Null
        TBGravar.Update
    End If
    USMsgBox ("Cadastro excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/Solicitação de ação"
    Evento = "Excluir outros"
    ID_documento = txtId
    Documento = "Nº SA: " & txtId
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos2
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 13, True
ProcCarregaToolBar2 Me, 15195, 10, True
ProcCarregaToolBar3 Me, 15195, 9, True
Formulario = "Qualidade/Solicitação de ação"
Direitos
Cmb_opcao_lista = "Validação"
SSTab1.Tab = 0
Permitido2 = False
ProcLimpaVariaveisPrincipais

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/RNC"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmCQ_SA_abrir.Show 1

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
Vendas_PI = False
Manutencao = False
Compras_Cotacao = False
Usuarios = False
Inspecao_recebimento = False
Funcionario = False
RNC = False
SolicitacaoAcao = True
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

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) solicitação(ões) de ação?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from CQ_SA where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from CQ_SA_Equipe where id_SA = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Solicitação de ação"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº SA: " & .ListItems(InitFor)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) de ação antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Solicitação(ões) de ação excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    Novo_SA = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir1()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListView2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) responsável(is)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from CQ_SA_Equipe where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Solicitação de ação"
            Evento = "Excluir responsável"
            ID_documento = txtId
            Documento = "Nº SA: " & txtId
            Documento1 = "Responsável: " & .ListItems(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) responsável(is) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Responsável(is) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos1
    ProcCarregaLista1
    Novo_SA1 = False
    Frame12.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe a solicitação de ação antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    ProcFiltrar
    Exit Sub
End If
NomeRel = "CQ_SA.rpt"
ProcImprimirRel "{CQ_SA.id} = " & txtId, ""

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
optNao.Value = True
Novo_SA = True
Frame1.Enabled = True
cmbTipo.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo1()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "solicitação de ação", "responsável", False) = False Then Exit Sub

ProcLimpaCampos1
Novo_SA1 = True
Frame12.Enabled = True
Cmb_responsavel.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_SA = True Then
    If USMsgBox("A solicitação de ação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_SA = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
procAlterar_abas True

Novo_SA = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "solicitação de ação", "o dados", False) = False Then Exit Sub

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If cmbTipo = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo.SetFocus
    Exit Sub
End If

If txtObjetivo.Text = "" Then
    NomeCampo = "o objetivo"
    ProcVerificaAcao
    txtObjetivo.SetFocus
    Exit Sub
End If



Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_SA where ID = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!Data = Date
    TBGravar!Responsavel = pubUsuario
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "a mesma", "solicitação de ação", False) = False Then Exit Sub
End If
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
ProcCarregaLista (1)
If Novo_SA = True Then
    USMsgBox ("Nova solicitação de ação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And ListView1.ListItems.Count <> 0 Then
        ListView1.SelectedItem = ListView1.ListItems(CodigoLista)
        ListView1.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Solicitação de ação"
ID_documento = txtId
Documento = "Nº SA: " & txtId
Documento1 = ""
ProcGravaEvento
'==================================
Novo_SA = False

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
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "solicitação de ação", "o responsável", False) = False Then Exit Sub

If Frame12.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_responsavel = "" Then
    NomeCampo = "o responsável"
    ProcVerificaAcao
    Cmb_responsavel.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CQ_SA_Equipe where ID = " & IIf(txtId1 = "", 0, txtId1), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_SA = txtId
TBGravar!Responsavel = Cmb_responsavel
TBGravar!Setor = Txt_setor
TBGravar.Update
txtId1 = TBGravar!ID
TBGravar.Close
ProcCarregaLista1
If Novo_SA1 = True Then
    USMsgBox ("Novo responsável cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo responsável"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar responsável"
    If CodigoLista1 <> 0 And ListView2.ListItems.Count <> 0 Then
        ListView2.SelectedItem = ListView2.ListItems(CodigoLista1)
        ListView2.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Solicitação de ação"
Documento = txtId.Text
Documento = "Nº SA: " & txtId
Documento1 = "Responsável: " & Cmb_responsavel
ProcGravaEvento
'==================================
Novo_SA1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("CQ_SA", "ID = " & .ListItems(InitFor), True) = False Then GoTo Proximo
                End If
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("CQ_SA", "ID = " & .ListItems(InitFor), "mesma", "solicitação de ação", "excluir esta", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
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

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_SA where id = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = ListView1.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListView2
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("CQ_SA", "ID = " & txtId, True) = False Then GoTo Proximo
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListView2, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("CQ_SA", "ID = " & txtId, "solicitação de ação", "responsável", "excluir este", False, True) = False Then
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

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListView2.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos1
txtId1 = ListView2.SelectedItem
Cmb_responsavel = ListView2.SelectedItem.ListSubItems(1)
1:
    Txt_setor = ListView2.SelectedItem.ListSubItems(2)
    CodigoLista1 = ListView2.SelectedItem.index
    Novo_SA1 = False
    Frame12.Enabled = True

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o responsável dessa solicitação de ação."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optNao_Click()
On Error GoTo tratar_erro

Frame10.Enabled = False
txtData_revisar = "__/__/____"
txtResponsavel_revisar = ""
txtTexto6 = ""
CQ_SA_Texto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSim_Click()
On Error GoTo tratar_erro

Frame10.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        procAlterar_abas False
        ListView1.SetFocus
    Case 1:
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        procAlterar_abas False
        ListView2.SetFocus
        'Carrega responsável
        Cmb_responsavel.Clear
        Set TBUsuarios = CreateObject("adodb.recordset")
        TBUsuarios.Open "Select Nome from Usuarios where Nome <> 'Null' order by Nome", Conexao, adOpenKeyset, adLockOptimistic
        If TBUsuarios.EOF = False Then
            Do While TBUsuarios.EOF = False
                Cmb_responsavel.AddItem TBUsuarios!Nome
                TBUsuarios.MoveNext
            Loop
        End If
        TBUsuarios.Close
        ProcLimpaCampos1
        ProcCarregaLista1
    Case 2:
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        SSTab2.SetFocus
        ProcLimpaCampos2
        ProcPuxadados2
        SSTab2.Tab = 0
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_SA = True Then
    USMsgBox ("Salve a solicitação de ação antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Procbloqueia()
On Error GoTo tratar_erro

Framelista.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDesbloqueia()
On Error GoTo tratar_erro

Framelista.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

If StrSql_CQ_SA_Localizar = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
Set TBLISTA_CQ_SA = CreateObject("adodb.recordset")
TBLISTA_CQ_SA.Open StrSql_CQ_SA_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_CQ_SA.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLISTA_CQ_SA.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_CQ_SA.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_CQ_SA.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_CQ_SA.RecordCount - IIf(Pagina > 1, (TBLISTA_CQ_SA.PageSize * (Pagina - 1)), 0), TBLISTA_CQ_SA.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_CQ_SA.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLISTA_CQ_SA!ID
        .Item(.Count).SubItems(1) = TBLISTA_CQ_SA!ID
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_CQ_SA!RNC), "", TBLISTA_CQ_SA!RNC)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_CQ_SA!Data), "", Format(TBLISTA_CQ_SA!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_CQ_SA!Responsavel), "", TBLISTA_CQ_SA!Responsavel)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_CQ_SA!Tipo), "", TBLISTA_CQ_SA!Tipo)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_CQ_SA!DtValidacao) = False, "Sim", "Não")
    End With
    TBLISTA_CQ_SA.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_CQ_SA.RecordCount
If TBLISTA_CQ_SA.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_CQ_SA.PageCount
ElseIf TBLISTA_CQ_SA.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_CQ_SA.PageCount & " de: " & TBLISTA_CQ_SA.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_CQ_SA.AbsolutePage - 1 & " de: " & TBLISTA_CQ_SA.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista1()
On Error GoTo tratar_erro

ListView2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CQ_SA_Equipe where ID_SA = " & txtId & " order by Responsavel", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListView2.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!Tipo = cmbTipo
TBGravar!Objetivo = txtObjetivo
TBGravar!Setor = txtSetor
TBGravar!eficaz = IIf(cmbEficaz = "", Null, cmbEficaz)
TBGravar!FIM = IIf(txtfim = "__/__/____", Null, Format(txtfim, "dd/mm/yyyy"))
TBGravar!Previsao = txtPrev
TBGravar!status = IIf(cmbStatus = "", Null, cmbStatus)
TBGravar!Obs = txtObs
TBGravar!RNC = txtRNC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados2()
On Error GoTo tratar_erro

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from CQ_SA where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtTexto1 = IIf(IsNull(TBFornecedor!Texto1), "", TBFornecedor!Texto1)
    txtTexto2 = IIf(IsNull(TBFornecedor!Texto2), "", TBFornecedor!Texto2)
    txtTexto3 = IIf(IsNull(TBFornecedor!Texto3), "", TBFornecedor!Texto3)
    txtTexto4 = IIf(IsNull(TBFornecedor!texto4), "", TBFornecedor!texto4)
    txtTexto5 = IIf(IsNull(TBFornecedor!Texto5), "", TBFornecedor!Texto5)
    txtTexto6 = IIf(IsNull(TBFornecedor!Texto6), "", TBFornecedor!Texto6)
    txtRiscos = IIf(IsNull(TBFornecedor!Riscos), "", TBFornecedor!Riscos)
    txtData1 = IIf(IsNull(TBFornecedor!Data1), "__/__/____", Format(TBFornecedor!Data1, "dd/mm/yyyy"))
    txtData2 = IIf(IsNull(TBFornecedor!Data2), "__/__/____", Format(TBFornecedor!Data2, "dd/mm/yyyy"))
    txtData3 = IIf(IsNull(TBFornecedor!Data3), "__/__/____", Format(TBFornecedor!Data3, "dd/mm/yyyy"))
    txtData4 = IIf(IsNull(TBFornecedor!Data4), "__/__/____", Format(TBFornecedor!Data4, "dd/mm/yyyy"))
    txtData5 = IIf(IsNull(TBFornecedor!Data5), "__/__/____", Format(TBFornecedor!Data5, "dd/mm/yyyy"))
    txtData6 = IIf(IsNull(TBFornecedor!Data6), "__/__/____", Format(TBFornecedor!Data6, "dd/mm/yyyy"))
    txtdata7 = IIf(IsNull(TBFornecedor!Data7), "__/__/____", Format(TBFornecedor!Data7, "dd/mm/yyyy"))
    txtData_revisar = IIf(IsNull(TBFornecedor!Data_revisar), "__/__/____", Format(TBFornecedor!Data_revisar, "dd/mm/yyyy"))
    txtData_riscos = IIf(IsNull(TBFornecedor!Data_riscos), "__/__/____", Format(TBFornecedor!Data_riscos, "dd/mm/yyyy"))
    txtResponsavel1 = IIf(IsNull(TBFornecedor!Responsavel1), "", TBFornecedor!Responsavel1)
    txtResponsavel2 = IIf(IsNull(TBFornecedor!responsavel2), "", TBFornecedor!responsavel2)
    txtResponsavel3 = IIf(IsNull(TBFornecedor!responsavel3), "", TBFornecedor!responsavel3)
    txtResponsavel4 = IIf(IsNull(TBFornecedor!responsavel4), "", TBFornecedor!responsavel4)
    txtResponsavel5 = IIf(IsNull(TBFornecedor!responsavel5), "", TBFornecedor!responsavel5)
    txtResponsavel_revisar = IIf(IsNull(TBFornecedor!Responsavel_revisar), "", TBFornecedor!Responsavel_revisar)
    txtResponsavel_riscos = IIf(IsNull(TBFornecedor!Responsavel_riscos), "", TBFornecedor!Responsavel_riscos)
    If TBFornecedor!Revisar = True Then optSim.Value = True Else optNao.Value = True
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

procAlterar_abas False
If CQ_SA_Alteracao = False Then
    ProcLimpaCampos2
    ProcPuxadados2
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

Private Sub txtRiscos_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtRiscos
CQ_SA_Aba = 17

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto1_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtTexto1
CQ_SA_Aba = 3

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttexto2_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtTexto2
CQ_SA_Aba = 6

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto3_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtTexto3
CQ_SA_Aba = 10

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto4_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtTexto4
CQ_SA_Aba = 14

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto5_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtTexto5
CQ_SA_Aba = 20

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto6_Click()
On Error GoTo tratar_erro

CQ_SA_Texto = txtTexto6
CQ_SA_Aba = 23

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
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: procFiltrarTodos
    Case 9: ProcValidarRegistros ListView1, "Qualidade/Solicitação de ação"
    'Case 11: ProcAjuda
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
    Case 1: ProcNovo1
    Case 2: procSalvar1
    Case 3: ProcExcluir1
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    'Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar2
    Case 2: procExcluir2
    Case 3: ProcImprimir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAlterar_abas(CQ_SA_sair As Boolean)
On Error GoTo tratar_erro

If txtDtValidacao <> "" Then Exit Sub
If CQ_SA_Aba = 1 Or CQ_SA_Aba = 2 Or CQ_SA_Aba = 3 Then
    If CQ_SA_Aba = 1 And CQ_SA_Texto <> txtData1 Then CQ_SA_Alteracao = True
    If CQ_SA_Aba = 2 And CQ_SA_Texto <> txtResponsavel1 Then CQ_SA_Alteracao = True
    If CQ_SA_Aba = 3 And CQ_SA_Texto <> txtTexto1 Then CQ_SA_Alteracao = True
    Texto_alterarAba = "A descrição da N/C ainda não foi salva"
ElseIf CQ_SA_Aba = 4 Or CQ_SA_Aba = 5 Or CQ_SA_Aba = 6 Then
        If CQ_SA_Aba = 4 And CQ_SA_Texto <> txtData2 Then CQ_SA_Alteracao = True
        If CQ_SA_Aba = 5 And CQ_SA_Texto <> txtResponsavel2 Then CQ_SA_Alteracao = True
        If CQ_SA_Aba = 6 And CQ_SA_Texto <> txtTexto2 Then CQ_SA_Alteracao = True
        Texto_alterarAba = "A causa/raiz da N/C ainda não foi salva"
    ElseIf CQ_SA_Aba = 7 Or CQ_SA_Aba = 8 Or CQ_SA_Aba = 9 Or CQ_SA_Aba = 10 Then
            If CQ_SA_Aba = 7 And CQ_SA_Texto <> txtData3 Then CQ_SA_Alteracao = True
            If CQ_SA_Aba = 8 And CQ_SA_Texto <> txtResponsavel3 Then CQ_SA_Alteracao = True
            If CQ_SA_Aba = 9 And CQ_SA_Texto <> txtData6 Then CQ_SA_Alteracao = True
            If CQ_SA_Aba = 10 And CQ_SA_Texto <> txtTexto3 Then CQ_SA_Alteracao = True
            Texto_alterarAba = "A ação corretiva ainda não foi salva"
        ElseIf CQ_SA_Aba = 11 Or CQ_SA_Aba = 12 Or CQ_SA_Aba = 13 Or CQ_SA_Aba = 14 Then
                If CQ_SA_Aba = 11 And CQ_SA_Texto <> txtData4 Then CQ_SA_Alteracao = True
                If CQ_SA_Aba = 12 And CQ_SA_Texto <> txtResponsavel4 Then CQ_SA_Alteracao = True
                If CQ_SA_Aba = 13 And CQ_SA_Texto <> txtdata7 Then CQ_SA_Alteracao = True
                If CQ_SA_Aba = 14 And CQ_SA_Texto <> txtTexto4 Then CQ_SA_Alteracao = True
                Texto_alterarAba = "O acompanhamento ainda não foi salvo"
            ElseIf CQ_SA_Aba = 15 Or CQ_SA_Aba = 16 Or CQ_SA_Aba = 17 Then
                    If CQ_SA_Aba = 15 And CQ_SA_Texto <> txtData_riscos Then CQ_SA_Alteracao = True
                    If CQ_SA_Aba = 16 And CQ_SA_Texto <> txtResponsavel_riscos Then CQ_SA_Alteracao = True
                    If CQ_SA_Aba = 17 And CQ_SA_Texto <> txtRiscos Then CQ_SA_Alteracao = True
                    Texto_alterarAba = "Os riscos e oportunidades ainda não foram salvos"
                ElseIf CQ_SA_Aba = 18 Or CQ_SA_Aba = 19 Or CQ_SA_Aba = 20 Then
                        If CQ_SA_Aba = 18 And CQ_SA_Texto <> txtData5 Then CQ_SA_Alteracao = True
                        If CQ_SA_Aba = 19 And CQ_SA_Texto <> txtResponsavel5 Then CQ_SA_Alteracao = True
                        If CQ_SA_Aba = 20 And CQ_SA_Texto <> txtTexto5 Then CQ_SA_Alteracao = True
                        Texto_alterarAba = "O fechamento ainda não foi salvo"
                    ElseIf CQ_SA_Aba = 21 Or CQ_SA_Aba = 22 Or CQ_SA_Aba = 23 Then
                        If CQ_SA_Aba = 21 And CQ_SA_Texto <> txtData_revisar Then CQ_SA_Alteracao = True
                        If CQ_SA_Aba = 22 And CQ_SA_Texto <> txtResponsavel_revisar Then CQ_SA_Alteracao = True
                        If CQ_SA_Aba = 23 And CQ_SA_Texto <> txtTexto6 Then CQ_SA_Alteracao = True
                        Texto_alterarAba = "Revisar documentos ainda não foi salvo"
End If

If CQ_SA_Alteracao = True Then
    If CQ_SA_sair = True Then
        If USMsgBox(Texto_alterarAba & ", deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
            procSalvar2
            If CQ_SA_Alteracao = True Then
                Exit Sub
            Else
                Unload Me
            End If
        End If
    Else
        If USMsgBox(Texto_alterarAba & ", deseja salvar antes de continuar?", vbYesNo) = vbYes Then procSalvar2
    End If
End If
CQ_SA_Alteracao = False
CQ_SA_Aba = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
