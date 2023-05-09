VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFinanceiro_familia 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Financeiro - Plano de contas"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Begin MSComctlLib.ListView Lista 
      Height          =   7045
      Left            =   75
      TabIndex        =   0
      Top             =   2660
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12435
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
      BorderStyle     =   1
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
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   15002
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Visualização"
      TabPicture(0)   =   "frmFinanceiro_familia.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "USTreeView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Gerenciamento"
      TabPicture(1)   =   "frmFinanceiro_familia.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "PBLista"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   825
         Left            =   75
         TabIndex        =   14
         Top             =   1815
         Width           =   15195
         Begin VB.TextBox txtIDFamilia 
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
            Left            =   7140
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Id."
            Top             =   360
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox Txt_descricao 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6420
            MaxLength       =   100
            TabIndex        =   17
            ToolTipText     =   "Descrição."
            Top             =   375
            Width           =   8595
         End
         Begin VB.TextBox TxtResponsavel 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3195
         End
         Begin VB.TextBox txtdata 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1125
         End
         Begin MSMask.MaskEdBox Txt_codigo 
            Height          =   315
            Left            =   4530
            TabIndex        =   19
            ToolTipText     =   "Código."
            Top             =   375
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#.##.##.##.##.##.##.##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código*"
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
            Left            =   5175
            TabIndex        =   23
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição*"
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
            Left            =   10327
            TabIndex        =   22
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Index           =   11
            Left            =   2460
            TabIndex        =   21
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   10
            Left            =   570
            TabIndex        =   20
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   495
         Left            =   75
         TabIndex        =   8
         Top             =   1320
         Width           =   11025
         Begin VB.CheckBox Chk_multa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Conta de multa"
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
            Left            =   5160
            TabIndex        =   12
            Top             =   180
            Width           =   2085
         End
         Begin VB.CheckBox Chk_juros 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Conta de juros"
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
            Left            =   3600
            TabIndex        =   11
            Top             =   180
            Width           =   2055
         End
         Begin VB.OptionButton optReceber 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Contas a receber*"
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
            Left            =   1770
            TabIndex        =   10
            Top             =   180
            Width           =   1665
         End
         Begin VB.OptionButton optPagar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Contas a pagar*"
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
            TabIndex        =   9
            Top             =   180
            Width           =   1515
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11100
         TabIndex        =   4
         Top             =   1320
         Width           =   4155
         Begin VB.OptionButton Opt_receber 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Contas a receber"
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
            Left            =   1650
            TabIndex        =   7
            Top             =   210
            Width           =   1575
         End
         Begin VB.OptionButton Opt_pagar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Contas a pagar"
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
            TabIndex        =   6
            Top             =   210
            Width           =   1425
         End
         Begin VB.OptionButton Opt_todos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Todos"
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
            Left            =   3270
            TabIndex        =   5
            Top             =   210
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74940
         TabIndex        =   2
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   3
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Ajuda"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Ajuda (F1)"
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
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Sair"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Sair (Esc)"
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
         ButtonWidth2    =   26
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonKey3      =   "11"
         ButtonAlignment3=   2
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   5
         ButtonLeft3     =   68
         ButtonTop3      =   2
         ButtonWidth3    =   24
         ButtonHeight3   =   24
         ButtonUseMaskColor3=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13770
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFinanceiro_familia.frx":0038
            Count           =   1
         End
      End
      Begin DrawSuite2022.USTreeView USTreeView1 
         Height          =   8655
         Left            =   -74940
         TabIndex        =   3
         Top             =   1290
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   15266
         BorderColor     =   12500670
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   13
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   115
         ButtonTop4      =   2
         ButtonWidth4    =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   51
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   209
         ButtonTop6      =   2
         ButtonWidth6    =   47
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   258
         ButtonTop7      =   2
         ButtonWidth7    =   46
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Atualizar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Utilizado pelo administrador do sistema."
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
         ButtonWidth8    =   50
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "De, para"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Atualizar de, para."
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
         ButtonLeft9     =   358
         ButtonTop9      =   2
         ButtonWidth9    =   50
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
         ButtonLeft10    =   410
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   414
         ButtonTop11     =   2
         ButtonWidth11   =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   452
         ButtonTop12     =   2
         ButtonWidth12   =   26
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
         ButtonLeft13    =   480
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   9330
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFinanceiro_familia.frx":1661
            Count           =   1
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   9690
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
   End
End
Attribute VB_Name = "frmFinanceiro_familia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Plano_Contas As Boolean 'OK
Public SQL_Plano_Contas As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=JYzY1WI_qqY&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=38&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362P" Then frmFinanceiro_familia_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmFinanceiro_familia_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza plano de contas em todos os módulos
            ProcRotinaAtualizacao "projproduto", "Subfamilia_financeiro"
            ProcRotinaAtualizacao "projfamilia", "Subfamilia_financeiro"
            ProcRotinaAtualizacao "Familia_financeiro", "Subfamilia"
            
            'Cria primeiro nível a pagar
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from tbl_familia where Codigo = '2.00.00.00.00.00.00.00'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then TBGravar.AddNew
            TBGravar!Destino = "P"
            TBGravar!Data = Date
            TBGravar!Responsavel = pubUsuario
            TBGravar!CODIGO = "2.00.00.00.00.00.00.00"
            TBGravar!Txt_descricao = "COMPRAS"
            TBGravar!Nivel = 1
            TBGravar.Update
            
            'Coloca o código nas contas (família)
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from tbl_familia where Tipo = 'F' and Destino = 'P' order by txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                TBGravar.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBGravar.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBGravar.MoveFirst
                Contador2 = 1
                Do While TBGravar.EOF = False
                    If Contador2 < 10 Then
                        Familiatext = "2.0" & Contador2 & ".00.00.00.00.00.00"
                    Else
                        Familiatext = "2." & Contador2 & ".00.00.00.00.00.00"
                    End If
                    TBGravar!CODIGO = Familiatext
                    TBGravar!Nivel = VerifNivel(Familiatext)
                    TBGravar.Update
                    
                    'Coloca o código nas contas (subfamília)
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_familia where ID_relacionamento = " & TBGravar!int_codfamilia & " and Destino = 'P' order by txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Qtde = 1
                        Do While TBAbrir.EOF = False
                            If Contador2 < 10 Then
                                If Qtde < 10 Then
                                    Familiatext = "2.0" & Contador2 & ".0" & Qtde & ".00.00.00.00.00"
                                Else
                                    Familiatext = "2.0" & Contador2 & "." & Qtde & ".00.00.00.00.00"
                                End If
                            Else
                                If Qtde < 10 Then
                                    Familiatext = "2." & Contador2 & ".0" & Qtde & ".00.00.00.00.00"
                                Else
                                    Familiatext = "2." & Contador2 & "." & Qtde & ".00.00.00.00.00"
                                End If
                            End If
                            TBAbrir!CODIGO = Familiatext
                            TBAbrir!Nivel = VerifNivel(Familiatext)
                            TBAbrir.Update
                            Qtde = Qtde + 1
                            TBAbrir.MoveNext
                        Loop
                    End If
                    TBAbrir.Close
                        
                    Contador = Contador + 1
                    Contador2 = Contador2 + 1
                    PBLista.Value = Contador
                    TBGravar.MoveNext
                Loop
            End If
                    
            'Cria primeiro nível a receber
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from tbl_familia where Codigo = '1.00.00.00.00.00.00.00'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then TBGravar.AddNew
            TBGravar!Destino = "R"
            TBGravar!Data = Date
            TBGravar!Responsavel = pubUsuario
            TBGravar!CODIGO = "1.00.00.00.00.00.00.00"
            TBGravar!Txt_descricao = "CLIENTES - OUTROS"
            TBGravar!Nivel = 1
            TBGravar.Update
            
            'Coloca o código nas contas (família)
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from tbl_familia where Tipo = 'F' and Destino = 'R' order by txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                TBGravar.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBGravar.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBGravar.MoveFirst
                Contador2 = 1
                Do While TBGravar.EOF = False
                    If Contador2 < 10 Then
                        Familiatext = "1.0" & Contador2 & ".00.00.00.00.00.00"
                    Else
                        Familiatext = "1." & Contador2 & ".00.00.00.00.00.00"
                    End If
                    TBGravar!CODIGO = Familiatext
                    TBGravar!Nivel = VerifNivel(Familiatext)
                    TBGravar.Update
                    
                    'Coloca o código nas contas (subfamília)
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_familia where ID_relacionamento = " & TBGravar!int_codfamilia & " and Tipo = 'SF' and Destino = 'R' order by txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Qtde = 1
                        Do While TBAbrir.EOF = False
                            If Contador2 < 10 Then
                                If Qtde < 10 Then
                                    Familiatext = "1.0" & Contador2 & ".0" & Qtde & ".00.00.00.00.00"
                                Else
                                    Familiatext = "1.0" & Contador2 & "." & Qtde & ".00.00.00.00.00"
                                End If
                            Else
                                If Qtde < 10 Then
                                    Familiatext = "1." & Contador2 & ".0" & Qtde & ".00.00.00.00.00"
                                Else
                                    Familiatext = "1." & Contador2 & "." & Qtde & ".00.00.00.00.00"
                                End If
                            End If
                            TBAbrir!CODIGO = Familiatext
                            TBAbrir!Nivel = VerifNivel(Familiatext)
                            TBAbrir.Update
                            Qtde = Qtde + 1
                            TBAbrir.MoveNext
                        Loop
                    End If
                    TBAbrir.Close
                        
                    Contador = Contador + 1
                    Contador2 = Contador2 + 1
                    PBLista.Value = Contador
                    TBGravar.MoveNext
                Loop
            End If
            TBGravar.Close
        End If
        
        If .Chk2.Value = 1 Then
            'Atualiza conta contábil nas contas
            Set TBFamilia = CreateObject("adodb.recordset")
            TBFamilia.Open "Select * from Familia_financeiro where ID_PC is not null order by IDConta, ID_PC", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then
                TBFamilia.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBFamilia.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBFamilia.MoveFirst
                Do While TBFamilia.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select Sum(Valor) as Valor from Familia_financeiro where IDConta = " & TBFamilia!IDConta & " and ID_PC = " & TBFamilia!ID_PC & " and TipoConta = '" & TBFamilia!TipoConta & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                    End If
                    TBAbrir.Close
                    
                    TBFamilia!valor = Format(valor, "###,##0.00")
                    TBFamilia.Update
                    
                    Conexao.Execute "DELETE from Familia_financeiro where ID <> " & TBFamilia!ID & " and IDConta = " & TBFamilia!IDConta & " and ID_PC = " & TBFamilia!ID_PC & " and TipoConta = '" & TBFamilia!TipoConta & "' and Deposito_transf = 'False'"
                    
                    Contador = Contador + 1
                    PBLista.Value = Contador
                    TBFamilia.MoveNext
                Loop
            End If
            TBFamilia.Close
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Financeiro/Plano de contas"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
        
        ProcCarregaListaPlano
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRotinaAtualizacao(NomeTabela As String, NomeCampo As String)
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "select * from " & NomeTabela & " where " & NomeCampo & " is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBGravar.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBGravar.MoveFirst
    Do While TBGravar.EOF = False
        If NomeTabela = "Familia_financeiro" Then NomeCampo1 = IIf(IsNull(TBGravar!SubFamilia), "", TBGravar!SubFamilia) Else NomeCampo1 = IIf(IsNull(TBGravar!Subfamilia_financeiro), "", TBGravar!Subfamilia_financeiro)
        If NomeCampo1 <> "" Then
            Set TBFamilia = CreateObject("adodb.recordset")
            TBFamilia.Open "select * from tbl_familia where txt_descricao = '" & NomeCampo1 & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then
                TBGravar!ID_PC = TBFamilia!int_codfamilia
                TBGravar.Update
            End If
            TBFamilia.Close
        End If
        Contador = Contador + 1
        PBLista.Value = Contador
        TBGravar.MoveNext
    Loop
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDFamilia = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_familia order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("int_codfamilia = " & txtIDFamilia)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtIDFamilia = TBLISTA!int_codfamilia
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from tbl_familia where int_codfamilia = " & txtIDFamilia, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDadosPlano
    Else
        USMsgBox ("Fim dos cadastros de contas do plano."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDFamilia = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_familia order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("int_codfamilia = " & txtIDFamilia)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtIDFamilia = TBLISTA!int_codfamilia
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from tbl_familia where int_codfamilia = " & txtIDFamilia, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDadosPlano
    Else
        USMsgBox ("Fim dos cadastros de contas do plano."), vbInformation, "CAPRIND v5.0"
    End If
End If

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
            Case vbKeyF3: ProcGravar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
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

Sub ProcGravar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If optPagar.Value = False And optreceber.Value = False Then
    NomeCampo = "o destino"
    ProcVerificaAcao
    Exit Sub
End If
If IsNumeric(txt_Codigo) = False Then
    NomeCampo = "o código"
    ProcVerificaAcao
    txt_Codigo.SetFocus
    Exit Sub
End If
If Txt_descricao.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao.SetFocus
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select int_codfamilia from tbl_familia where int_codfamilia <> " & txtIDFamilia & " and Codigo = '" & txt_Codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Este código está sendo utilizado, favor alterar."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    txt_Codigo.SetFocus
    Exit Sub
End If
TBAbrir.Close

If optPagar.Value = True Then DestinoTexto = "P" Else DestinoTexto = "R"
Permitido = True
If Chk_juros.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Codigo, txt_descricao, Juros from tbl_familia where int_codfamilia <> " & txtIDFamilia & " and Juros = 'True' and Destino = '" & DestinoTexto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If USMsgBox("A conta " & TBAbrir!CODIGO & " - " & TBAbrir!Txt_descricao & " já está classificada como juros, deseja alterar?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            TBAbrir!Juros = False
            TBAbrir.Update
        Else
            Permitido = False
        End If
    End If
    TBAbrir.Close
End If
If Chk_multa.Value = 1 Then
    Permitido1 = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Codigo, txt_descricao, Multa from tbl_familia where int_codfamilia <> " & txtIDFamilia & " and Multa = 'True' and Destino = '" & DestinoTexto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If USMsgBox("A conta " & TBAbrir!CODIGO & " - " & TBAbrir!Txt_descricao & " já está classificada como multa, deseja alterar?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        TBAbrir!Multa = False
            TBAbrir.Update
        Else
            Permitido1 = False
        End If
    End If
    TBAbrir.Close
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_familia where int_codfamilia = " & txtIDFamilia, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    versao = "v" & App.Major & "." & App.Minor & "." & App.Revision
    If versao <= "v4.5.045" Then
        Conexao.Execute "Update Familia_financeiro Set Familia = '" & Txt_descricao & "' where Familia = '" & TBGravar!Txt_descricao & "'"
        Conexao.Execute "Update Familia_financeiro Set Subfamilia = '" & Txt_descricao & "' where Familia = '" & TBGravar!Txt_descricao & "'"
        Conexao.Execute "Update projproduto Set Familia_financeiro = '" & Txt_descricao & "' where Familia_financeiro = '" & TBGravar!Txt_descricao & "'"
        Conexao.Execute "Update projproduto Set Subfamilia_financeiro = '" & Txt_descricao & "' where Subfamilia_financeiro = '" & TBGravar!Txt_descricao & "'"
    End If
End If
TBGravar!Destino = DestinoTexto
If Permitido = True Then If Chk_juros.Value = 1 Then TBGravar!Juros = True Else TBGravar!Juros = False
If Permitido1 = True Then If Chk_multa.Value = 1 Then TBGravar!Multa = True Else TBGravar!Multa = False
If txtData = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData
If txtResponsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel
TBGravar!CODIGO = txt_Codigo
TBGravar!Txt_descricao = Txt_descricao.Text
TBGravar!Nivel = VerifNivel(txt_Codigo)
TBGravar.Update
txtIDFamilia = TBGravar!int_codfamilia
TBGravar.Close

ProcCarregaListaPlano
If Novo_Plano_Contas = True Then
    USMsgBox ("Nova conta do plano cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 Then
        If Lista.ListItems.Count = 0 Then Exit Sub
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Financeiro/Plano de contas"
ID_documento = txtIDFamilia
Documento = "Código: " & txt_Codigo & " - Descrição: " & Txt_descricao
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Plano_Contas = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function VerifNivel(CODIGO As String) As Integer
On Error GoTo tratar_erro

Nivel8A = Mid(CODIGO, 21, 2)
Nivel7A = Mid(CODIGO, 18, 2)
Nivel6A = Mid(CODIGO, 15, 2)
Nivel5A = Mid(CODIGO, 12, 2)
Nivel4A = Mid(CODIGO, 9, 2)
Nivel3A = Mid(CODIGO, 6, 2)
Nivel2A = Mid(CODIGO, 3, 2)
            
If Nivel8A <> "00" Then
    VerifNivel = 8
ElseIf Nivel7A <> "00" Then
        VerifNivel = 7
    ElseIf Nivel6A <> "00" Then
            VerifNivel = 6
        ElseIf Nivel5A <> "00" Then
                VerifNivel = 5
            ElseIf Nivel4A <> "00" Then
                    VerifNivel = 4
                ElseIf Nivel3A <> "00" Then
                        VerifNivel = 3
                    ElseIf Nivel2A <> "00" Then
                            VerifNivel = 2
                        Else
                            VerifNivel = 1
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcCarregaListaPlano()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open SQL_Plano_Contas, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBAbrir.MoveFirst
    Do While TBAbrir.EOF = False
        With Lista.ListItems
            .Add , , IIf(IsNull(TBAbrir!int_codfamilia), "", TBAbrir!int_codfamilia)
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!CODIGO), "", TBAbrir!CODIGO)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
        End With
        TBAbrir.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtIDFamilia = 0
optPagar.Value = False
optreceber.Value = False
Chk_juros.Value = 0
Chk_multa.Value = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txt_Codigo = "_.__.__.__.__.__.__.__"
Txt_descricao.Text = ""
CodigoLista = 0
Caption = "Financeiro - Plano de contas"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Plano de contas"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Contas_plano de contas.rpt"
ProcImprimirRel "", ""
NomeRel = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmFinanceiro_familia_localizar.Show 1

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
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) conta(s) do plano?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_familia WHERE int_codfamilia = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Financeiro/Plano de contas"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then Documento = "Código: " & TBFI!CODIGO & " - Descrição: " & TBFI!Txt_descricao Else Documento = "Descrição: " & TBFI!Txt_descricao
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE from tbl_familia WHERE int_codfamilia = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) contas(s) do plano antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Conta(s) excluída(s) do plano com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaListaPlano
    Frame1.Enabled = False
    Frame8.Enabled = False
    Novo_Plano_Contas = False
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
ProcLimpaCampos
Novo_Plano_Contas = True
Frame1.Enabled = True
Frame8.Enabled = True
optPagar.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Plano_Contas = True Then
    If USMsgBox("A conta do plano ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        If Novo_Plano_Contas = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Plano_Contas = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Decrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosPlano()
On Error GoTo tratar_erro

txtIDFamilia = TBFamilia!int_codfamilia
If IsNull(TBFamilia!Destino) = False Then
    If TBFamilia!Destino = "P" Then optPagar.Value = True Else optreceber.Value = True
End If
If TBFamilia!Juros = True Then Chk_juros.Value = 1 Else Chk_juros.Value = 0
If TBFamilia!Multa = True Then Chk_multa.Value = 1 Else Chk_multa.Value = 0
txtData = IIf(IsNull(TBFamilia!Data), "", Format(TBFamilia!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBFamilia!Responsavel), "", TBFamilia!Responsavel)
txt_Codigo = IIf(IsNull(TBFamilia!CODIGO), "_.__.__.__.__.__.__.__", TBFamilia!CODIGO)
Txt_descricao.Text = IIf(IsNull(TBFamilia!Txt_descricao), "", TBFamilia!Txt_descricao)
Caption = "Financeiro - Plano de contas (Código : " & TBFamilia!CODIGO & " - Descrição : " & TBFamilia!Txt_descricao & ")"
Novo_Plano_Contas = False
Frame1.Enabled = True
Frame8.Enabled = True

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
                ProcVerificaRegistroUtilizadoSemMsg "projproduto", "ID_PC = " & .ListItems(InitFor)
                If Permitido = False Then GoTo Proximo
                ProcVerificaRegistroUtilizadoSemMsg "Familia_financeiro", "ID_PC = " & .ListItems(InitFor)
                If Permitido = False Then GoTo Proximo
                ProcVerificaRegistroUtilizadoSemMsg "Usuarios_Setor_Previsao", "ID_PC = " & .ListItems(InitFor)
                If Permitido = False Then GoTo Proximo
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
            Mensagem = "Não é permitido excluir esta conta do plano, pois a mesma está sendo utilizada no módulo"
            ProcVerificaRegistroUtilizado "projproduto", "ID_PC = " & .ListItems(InitFor), "Engenharia/Produtos e serviços"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If optPagar.Value = True Then Texto = "pagar" Else Texto = "receber"
            ProcVerificaRegistroUtilizado "Familia_financeiro", "ID_PC = " & .ListItems(InitFor), "Financeiro/Contas a " & Texto
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "Usuarios_Setor_Previsao", "ID_PC = " & .ListItems(InitFor), "Custos/Centro de custo"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDePara()
On Error GoTo tratar_erro

frmFinanceiro_familia_de_para.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from tbl_familia where int_codfamilia = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDadosPlano
    CodigoLista = Lista.SelectedItem.index
End If
TBFamilia.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 13, True
ProcCarregaToolBar2 Me, 15192, 3, True

Formulario = "Financeiro/Plano de contas"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0

SQL_Plano_Contas = "Select * from tbl_familia order by Codigo"

Lista.Visible = False
ProcCarregaVisualizacao

'ProcCarregaListaPlano

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_pagar_Click()
On Error GoTo tratar_erro

SQL_Plano_Contas = "Select * from tbl_familia where Destino = 'P' or Destino = 'PR' order by Codigo"
ProcCarregaListaPlano

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_receber_Click()
On Error GoTo tratar_erro

SQL_Plano_Contas = "Select * from tbl_familia where Destino = 'R' or Destino = 'PR' order by Codigo"
ProcCarregaListaPlano

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_todos_Click()
On Error GoTo tratar_erro

SQL_Plano_Contas = "Select * from tbl_familia where txt_descricao is not null order by Codigo"
ProcCarregaListaPlano

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 1:
        Lista.Visible = True
        If Lista.Visible = True Then Lista.SetFocus
    Case 0:
        Lista.Visible = False
        ProcCarregaVisualizacao
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaVisualizacao()
On Error GoTo tratar_erro

With USTreeView1
    .Clear
    
    'Adicionando as chaves principais
    Set Receber = .Nodes.AddNode("Contas a receber", "A", , True, , , , 0, vbBlue)
    Set Pagar = .Nodes.AddNode("Contas a pagar", "B", , True, , , , 0, vbRed)
    
    'Receber
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_familia where Codigo is not null order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            
            Descricao = TBAbrir!CODIGO & " - " & TBAbrir!Txt_descricao
            IDlista = TBAbrir!int_codfamilia
            Nivel = IIf(IsNull(TBAbrir!Nivel), 0, TBAbrir!Nivel)
           
            If Nivel = 8 Then
                If TBAbrir!Destino = "R" Then
                    .Nodes.AddNode Descricao, IDlista, , , , , , , , NivelR7
                Else
                    .Nodes.AddNode Descricao, IDlista, , , , , , , , NivelP7
                End If
            ElseIf Nivel = 7 Then
                    If TBAbrir!Destino = "R" Then
                        Set NivelR7 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR6)
                    Else
                        Set NivelP7 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP6)
                    End If
                ElseIf Nivel = 6 Then
                        If TBAbrir!Destino = "R" Then
                            Set NivelR6 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR5)
                        Else
                            Set NivelP6 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP5)
                        End If
                    ElseIf Nivel = 5 Then
                            If TBAbrir!Destino = "R" Then
                                Set NivelR5 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR4)
                            Else
                                Set NivelP5 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP4)
                            End If
                        ElseIf Nivel = 4 Then
                                If TBAbrir!Destino = "R" Then
                                    Set NivelR4 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR3)
                                Else
                                    Set NivelP4 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP3)
                                End If
                            ElseIf Nivel = 3 Then
                                    If TBAbrir!Destino = "R" Then
                                        Set NivelR3 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR2)
                                    Else
                                        Set NivelP3 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP2)
                                    End If
                                ElseIf Nivel = 2 Then
                                        If TBAbrir!Destino = "R" Then
                                            Set NivelR2 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelR1)
                                        Else
                                            Set NivelP2 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , NivelP1)
                                        End If
                                    Else
                                        If TBAbrir!Destino = "R" Then
                                            Set NivelR1 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , Receber)
                                        Else
                                            Set NivelP1 = .Nodes.AddNode(Descricao, IDlista, , , , , , , , Pagar)
                                        End If
            End If
            
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    .ExpandAllNodes False
End With

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
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcAtualizar
    Case 9: ProcDePara
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
    Case 1: ProcAjuda
    Case 2: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
