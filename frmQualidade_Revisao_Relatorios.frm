VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQualidade_Revisao_Relatorios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Histórico de revisões dos relatórios"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   11040
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmQualidade_Revisao_Relatorios.frx":0000
      Count           =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2550
      Top             =   5820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   30
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   15105
      _ExtentX        =   26644
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
      ButtonCaption5  =   "Atualizar"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Utilizado pelo administrador do sistema."
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
      ButtonWidth5    =   59
      ButtonHeight5   =   21
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
      ButtonLeft6     =   238
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
      ButtonLeft7     =   242
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
      ButtonLeft8     =   285
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
      ButtonLeft9     =   317
      ButtonTop9      =   2
      ButtonWidth9    =   24
      ButtonHeight9   =   24
      ButtonUseMaskColor9=   0   'False
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6945
      Left            =   90
      TabIndex        =   4
      Top             =   2160
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   12250
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
         Text            =   "Relatório"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
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
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Histórico"
         Object.Width           =   13079
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8985
      Left            =   60
      TabIndex        =   27
      Top             =   990
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   15849
      _Version        =   393216
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
      TabCaption(0)   =   "Dados dos relatórios"
      TabPicture(0)   =   "frmQualidade_Revisao_Relatorios.frx":4C79
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Revisões"
      TabPicture(1)   =   "frmQualidade_Revisao_Relatorios.frx":4C95
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "txtID1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Subreports"
      TabPicture(2)   =   "frmQualidade_Revisao_Relatorios.frx":4CB1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtID2"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "Lista2"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   30
         TabIndex        =   35
         Top             =   8130
         Width           =   15105
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
            TabIndex        =   6
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
            Left            =   3780
            TabIndex        =   5
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   10
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_Revisao_Relatorios.frx":4CCD
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
            TabIndex        =   9
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_Revisao_Relatorios.frx":8471
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
            TabIndex        =   7
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
            TabIndex        =   8
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_Revisao_Relatorios.frx":BF7A
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
            TabIndex        =   11
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmQualidade_Revisao_Relatorios.frx":10069
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
            Left            =   4410
            TabIndex        =   40
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
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label14 
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
            Left            =   3090
            TabIndex        =   36
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox txtID2 
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
         Left            =   -72000
         MaxLength       =   4
         MouseIcon       =   "frmQualidade_Revisao_Relatorios.frx":138F5
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Text            =   "0"
         ToolTipText     =   "ID"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1095
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
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   -74970
         TabIndex        =   32
         Top             =   330
         Width           =   15105
         Begin VB.TextBox Txt_nome_subreport 
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
            TabIndex        =   17
            ToolTipText     =   "Nome do relatório."
            Top             =   375
            Width           =   14745
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nome do subreport*"
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
            Left            =   6817
            TabIndex        =   33
            Top             =   180
            Width           =   1470
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -72030
         MaxLength       =   4
         MouseIcon       =   "frmQualidade_Revisao_Relatorios.frx":13BFF
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   "ID"
         Top             =   5010
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtID 
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
         Left            =   2970
         MaxLength       =   4
         MouseIcon       =   "frmQualidade_Revisao_Relatorios.frx":13F09
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Text            =   "0"
         ToolTipText     =   "ID"
         Top             =   4980
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame2 
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
         Height          =   1815
         Left            =   -74970
         TabIndex        =   22
         Top             =   330
         Width           =   15105
         Begin VB.TextBox Txt_historico 
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
            Height          =   765
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            ToolTipText     =   "Histórico da revisão."
            Top             =   915
            Width           =   14745
         End
         Begin VB.TextBox Txt_revisao 
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
            MaxLength       =   4
            TabIndex        =   12
            ToolTipText     =   "Número da revisão."
            Top             =   375
            Width           =   1095
         End
         Begin VB.TextBox Txt_responsavel 
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
            Left            =   2700
            MaxLength       =   50
            TabIndex        =   14
            ToolTipText     =   "Responsável pela revisão."
            Top             =   375
            Width           =   12225
         End
         Begin MSComCtl2.DTPicker Txt_data_revisao 
            Height          =   315
            Left            =   1290
            TabIndex        =   13
            ToolTipText     =   "Data da revisão."
            Top             =   375
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
            Format          =   199950337
            CurrentDate     =   39057
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. revisão"
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
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Histórico*"
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
            Left            =   7245
            TabIndex        =   25
            Top             =   735
            Width           =   705
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Revisão*"
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
            Left            =   397
            TabIndex        =   24
            Top             =   180
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pela revisão*"
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
            Left            =   7852
            TabIndex        =   23
            Top             =   180
            Width           =   1920
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
         Height          =   825
         Left            =   30
         TabIndex        =   18
         Top             =   330
         Width           =   15105
         Begin VB.CheckBox chkResponsavel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar responsável pela emissão"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11700
            TabIndex        =   39
            Top             =   270
            Width           =   3375
         End
         Begin VB.CheckBox Chk_personalizado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Personalizado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11700
            TabIndex        =   3
            Top             =   495
            Width           =   1785
         End
         Begin VB.TextBox Txt_nome_relatorio 
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
            Left            =   6990
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Nome do relatório."
            Top             =   375
            Width           =   4545
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   6540
            Picture         =   "frmQualidade_Revisao_Relatorios.frx":14213
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Localizar relatório."
            Top             =   375
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do relatório."
            Top             =   375
            Width           =   6345
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nome do relatório*"
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
            Left            =   8580
            TabIndex        =   21
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho do relatório*"
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
            Left            =   2565
            TabIndex        =   20
            Top             =   180
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView Lista1 
         Height          =   6585
         Left            =   -74970
         TabIndex        =   16
         Top             =   2160
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   11615
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Histórico"
            Object.Width           =   18371
         EndProperty
      End
      Begin MSComctlLib.ListView Lista2 
         Height          =   7575
         Left            =   -74970
         TabIndex        =   19
         Top             =   1170
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   13361
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
            Text            =   "Subreport"
            Object.Width           =   25426
         EndProperty
      End
   End
End
Attribute VB_Name = "frmQualidade_Revisao_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Revisao_Relatorios As Boolean 'OK
Dim Novo_Revisao_Relatorios1 As Boolean 'OK
Dim Novo_Revisao_Relatorios2 As Boolean 'OK
Public Sql_Revisao_Relatorios_Localizar As String 'OK
Dim TBLISTA_Revisao_Relatorios As ADODB.Recordset 'OK
Dim Arquivo As File 'OK
Dim Diretorio As Folder 'OK

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho
Txt_nome_relatorio = Nome_anexo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Revisao_Relatorios.AbsolutePage <> 2 Then
    If TBLISTA_Revisao_Relatorios.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Revisao_Relatorios.PageCount - 1)
    Else
        TBLISTA_Revisao_Relatorios.AbsolutePage = TBLISTA_Revisao_Relatorios.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Revisao_Relatorios.AbsolutePage)
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
    TBLISTA_Revisao_Relatorios.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Revisao_Relatorios.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Revisao_Relatorios.AbsolutePage = 1
ProcExibePagina (TBLISTA_Revisao_Relatorios.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Revisao_Relatorios.AbsolutePage <> -3 Then
    If TBLISTA_Revisao_Relatorios.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Revisao_Relatorios.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Revisao_Relatorios.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Revisao_Relatorios.AbsolutePage = TBLISTA_Revisao_Relatorios.PageCount
ProcExibePagina (TBLISTA_Revisao_Relatorios.AbsolutePage)

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
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: ProcSair
End Select

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
If Sql_Revisao_Relatorios_Localizar = "" Then Exit Sub
Set TBLISTA_Revisao_Relatorios = CreateObject("adodb.recordset")
TBLISTA_Revisao_Relatorios.Open Sql_Revisao_Relatorios_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Revisao_Relatorios.EOF = False Then
    ProcExibePagina (Pagina)
Else
    ProcLimpaCampos_Rel
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Revisao_Relatorios.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Revisao_Relatorios.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Revisao_Relatorios.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Revisao_Relatorios.RecordCount - IIf(Pagina > 1, (TBLISTA_Revisao_Relatorios.PageSize * (Pagina - 1)), 0), TBLISTA_Revisao_Relatorios.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Revisao_Relatorios.EOF = False And (ContadorReg <= TamanhoPagina)
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Qualidade_revisao_relatorios where Nome_relatorio = '" & TBLISTA_Revisao_Relatorios!Nome_relatorio & "' order by Revisao desc", Conexao, adOpenKeyset, adLockReadOnly
    If TBFI.EOF = False Then
        With Lista.ListItems
            .Add , , TBFI!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!Nome_relatorio), "", TBFI!Nome_relatorio)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBFI!Revisao), "", TBFI!Revisao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBFI!Data), "", Format(TBFI!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!Responsavel), "", TBFI!Responsavel)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBFI!Historico), "", TBFI!Historico)
        End With
    End If
    TBLISTA_Revisao_Relatorios.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Revisao_Relatorios.RecordCount
If TBLISTA_Revisao_Relatorios.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Revisao_Relatorios.PageCount
ElseIf TBLISTA_Revisao_Relatorios.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Revisao_Relatorios.PageCount & " de: " & TBLISTA_Revisao_Relatorios.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Revisao_Relatorios.AbsolutePage - 1 & " de: " & TBLISTA_Revisao_Relatorios.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaRevisao()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios where Nome_relatorio = '" & Txt_nome_relatorio & "' order by Data desc, revisao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    TBHistProc.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBHistProc.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBHistProc.MoveFirst
    Do While TBHistProc.EOF = False
        With Lista1.ListItems
            .Add , , TBHistProc!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBHistProc!Data), "", Format(TBHistProc!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBHistProc!Responsavel), "", TBHistProc!Responsavel)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBHistProc!Historico), "", TBHistProc!Historico)
        End With
        TBHistProc.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBHistProc.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaSubreport()
On Error GoTo tratar_erro

Lista2.ListItems.Clear
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios_subreports where Nome_relatorio = '" & Txt_nome_relatorio & "' order by Subreport", Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    TBHistProc.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBHistProc.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBHistProc.MoveFirst
    Do While TBHistProc.EOF = False
        With Lista2.ListItems
            .Add , , TBHistProc!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBHistProc!SubReport), "", TBHistProc!SubReport)
        End With
        TBHistProc.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBHistProc.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15200, 8, True
Formulario = "Qualidade/Histórico de revisão dos relatórios"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Txt_data_revisao.Value = Date

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Histórico de revisão dos relatórios"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0: ProcExcluir_Rel
    Case 1: ProcExcluir_Revisao
    Case 2: ProcExcluir_Subreport
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        ProcLimpaCampos_Rel
        PocLimparTudo
        Novo_Revisao_Relatorios = True
        Frame3.Enabled = True
        cmdImportar_Click
    Case 1:
        ProcLimpaCampos_Revisao
        Novo_Revisao_Relatorios1 = True
        Frame2.Enabled = True
        Txt_revisao.SetFocus
    Case 2:
        ProcLimpaCampos_Subreport
        Novo_Revisao_Relatorios2 = True
        Frame1.Enabled = True
        Txt_nome_subreport.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmQualidade_Revisao_Relatorios_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_Rel()
On Error GoTo tratar_erro

txtId = 0
txt_Caminho = ""
Txt_nome_relatorio = ""
Chk_personalizado.Value = 0
chkResponsavel.Value = 0
CodigoLista = 0
Caption = "Qualidade - Histórico de revisão dos relatórios"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_Revisao()
On Error GoTo tratar_erro

txtId1 = 0
Txt_revisao.Text = ""
Txt_data_revisao.Value = Date
Txt_responsavel.Text = ""
Txt_historico.Text = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_Subreport()
On Error GoTo tratar_erro

txtID2 = 0
Txt_nome_subreport = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Revisao_Relatorios = True Then
    If USMsgBox("O histórico ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Revisao_Relatorios = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Revisao_Relatorios1 = True Then
    If USMsgBox("A revisão ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Revisao_Relatorios1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Revisao_Relatorios2 = True Then
    If USMsgBox("O subreport ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Revisao_Relatorios2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Revisao_Relatorios = False
Novo_Revisao_Relatorios1 = False
Novo_Revisao_Relatorios2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0: ProcSalvar_Rel
    Case 1: ProcSalvar_Revisao
    Case 2: ProcSalvar_Subreport
End Select

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
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
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
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    ProcLimpaCampos_Rel
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBHistProc.Close
Frame3.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub PocLimparTudo()
On Error GoTo tratar_erro

ProcLimpaCampos_Revisao
ProcLimpaCampos_Subreport
Lista1.ListItems.Clear
Lista2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems(InitFor).ListSubItems(1) = 0 Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems(InitFor).ListSubItems(1) = 0 Then
                USMsgBox ("Não é permitido excluir a revisão 0."), vbExclamation, "CAPRIND v5.0"
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

Private Sub Lista1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista1.ListItems.Count = 0 Then Exit Sub
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios where id = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    ProcLimpaCampos_Revisao
    ProcCarregaDados_Revisao
    CodigoLista1 = Lista1.SelectedItem.index
End If
TBHistProc.Close
Frame2.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista2
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista2, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista2.ListItems.Count = 0 Then Exit Sub
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios_subreports where id = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    ProcLimpaCampos_Subreport
    ProcCarregaDados_Subreport
    CodigoLista2 = Lista2.SelectedItem.index
End If
TBHistProc.Close
Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId = TBHistProc!ID
Caption = "Qualidade - Histórico de revisão dos relatórios - (Nome do relatório :  " & IIf(IsNull(TBHistProc!Nome_relatorio), "", TBHistProc!Nome_relatorio) & " - Revisão : " & IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao) & ")"
txt_Caminho = IIf(IsNull(TBHistProc!caminho), "", TBHistProc!caminho)
Txt_nome_relatorio = IIf(IsNull(TBHistProc!Nome_relatorio), "", TBHistProc!Nome_relatorio)
If TBHistProc!Personalizado = True Then Chk_personalizado.Value = 1 Else Chk_personalizado.Value = 0
If TBHistProc!Responsavel_rel = True Then chkResponsavel.Value = 1 Else chkResponsavel.Value = 0
Novo_Revisao_Relatorios = False
PocLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados_Revisao()
On Error GoTo tratar_erro

txtId1 = TBHistProc!ID
Txt_revisao = IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao)
If IsNull(TBHistProc!Data) = False Then Txt_data_revisao.Value = TBHistProc!Data
Txt_responsavel = IIf(IsNull(TBHistProc!Responsavel), "", (TBHistProc!Responsavel))
Txt_historico = IIf(IsNull(TBHistProc!Historico), "", (TBHistProc!Historico))
Novo_Revisao_Relatorios1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados_Subreport()
On Error GoTo tratar_erro

txtID2 = TBHistProc!ID
Txt_nome_subreport = IIf(IsNull(TBHistProc!SubReport), "", TBHistProc!SubReport)
Novo_Revisao_Relatorios2 = False

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
Lista.Visible = True
With USToolBar1
    Select Case SSTab1.Tab
        Case 0:
            .ButtonState(2) = 0
            If Lista.Visible = True Then Lista.SetFocus
        Case 1:
            .ButtonState(2) = 5
            Lista.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            'Lista1.SetFocus
            ProcCarregaListaRevisao
        Case 2:
            .ButtonState(2) = 5
            Lista.Visible = False
            ProcVerificaProsseguir
            If Permitido = False Then Exit Sub
            'Lista2.SetFocus
            ProcCarregaListaSubreport
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Revisao_Relatorios = True Then
    USMsgBox ("Salve o relatório antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcAtualizar
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362Q" Then
    If USMsgBox("Deseja realmente atualizar os dados dos relatórios?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Conexao.Execute "Truncate Table Qualidade_revisao_relatorios"
        Conexao.Execute "DBCC CheckIdent('Qualidade_revisao_relatorios',Reseed,1)"
        Conexao.Execute "Truncate Table Qualidade_revisao_relatorios_subreports"
        Conexao.Execute "DBCC CheckIdent('Qualidade_revisao_relatorios_subreports',Reseed,1)"
        
        Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
        If GerArqPastas.FolderExists(Localrel) = True Then Call ProcVerifRelatorios(GerArqPastas.GetFolder(Localrel), False)
            
        Set GerArqPastas = CreateObject("Scripting.FileSystemObject")
        If GerArqPastas.FolderExists(Localrel & "\Personalizados") = True Then Call ProcVerifRelatorios(GerArqPastas.GetFolder(Localrel & "\Personalizados"), True)
         
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Histórico de revisão dos relatórios"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifRelatorios(Diretorio As Folder, Personalizado As Boolean)
On Error GoTo tratar_erro

For Each Arquivo In Diretorio.files
    If Arquivo.Name Like "*.rpt" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Qualidade_revisao_relatorios where Nome_relatorio = '" & Arquivo.Name & "' order by revisao desc", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBAbrir.AddNew
            TBAbrir!Revisao = 0
            TBAbrir!Data = Date
            TBAbrir!Responsavel = "PROCAM"
            If Personalizado = True Then TBAbrir!caminho = Localrel & "\Personalizados\" & Arquivo.Name Else TBAbrir!caminho = Localrel & "\" & Arquivo.Name
            TBAbrir!Nome_relatorio = Arquivo.Name
            TBAbrir!Personalizado = Personalizado
            TBAbrir.Update
            
            ProcGravarSubReports Arquivo.Name
        End If
    End If
Next

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarSubReports(Nome_relatorio As String)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Qualidade_revisao_relatorios where Nome_relatorio = '" & Nome_relatorio & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Select Case Nome_relatorio
        Case "Almoxarifado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Itens')"
        Case "Clientes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contatos')"
        Case "Clientes_lista_email.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contatos')"
        Case "Clientes_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_alteracoes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_cotacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_cotacao_fornecedores.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_cotacao_mapa.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "Compras_follow up de compras.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Qtde_recebida')"
        Case "Compras_historico_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_historico_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_historico_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_historico_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_indice_atraso_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_indice_atraso_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_indice_atraso_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_lista de pedidos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Compras_pedido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Listaprodutos'), ('" & TBAbrir!Nome_relatorio & "', 'Listaservicos')"
        Case "Compras_programacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Recebimento_pedido'), ('" & TBAbrir!Nome_relatorio & "', 'Programacaomes'), ('" & TBAbrir!Nome_relatorio & "', 'Programacaoprevista')"
        Case "Compras_requisicao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Pedidos.rpt')"
        Case "Contas_fluxodecaixa_projetado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Vencidos')"
        Case "Contas_fluxodecaixa_projetado_resumido_ano grafico": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_fluxodecaixa_projetado_resumido_dia grafico": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_fluxodecaixa_projetado_resumido_mes grafico": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_fluxodecaixa_projetado_saldos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Vencidos')"
        Case "Contas_fluxodecaixa_realizado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Vencidos')"
        Case "Contas_fluxodecaixa_realizado_resumido_ano grafico": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_fluxodecaixa_realizado_resumido_dia grafico": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_fluxodecaixa_realizado_resumido_mes grafico": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_fluxodecaixa_realizado_saldos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Vencidos')"
        Case "Contas_pagar.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagar_conta contabil.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagar_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagas_conta contabil.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagas_copia de cheque.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagas_copia de cheque_periodo.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_pagas_recibo.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Conta_contabil')"
        Case "Contas_pagas_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_plano de contas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber_bordero.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber_conta contabil.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber_descontada_local_desconto.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber_duplicatasoperacoes_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber_duplicatasoperacoes_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_receber_duplicata_selecionadas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'Referente')"
        Case "Contas_receber_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_recebidas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_recebidas_conta contabil.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_recebidas_descontada_instituicoes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_recebidas_descontada_local_desconto.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_recebidas_recomprada_local_desconto.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_recebidas_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_relatorio_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Contas_relatorio_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "Contas_relatorio_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "Contas_relatorio_razao_clientes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Movimentacao')"
        Case "Contas_relatorio_razao_fornecedores.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Movimentacao')"
        Case "Contas_relatorio_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "CQ_certificado_qualidade.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_inspecao recebimento.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_instrumentos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_nc.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_nc_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_nc_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_nc_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_nc_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_plano inspecao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Familias.rpt')"
        Case "CQ_plano inspecao_frequencia de medicao.rpt":  Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Familias')"
        Case "CQ_plano medicao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'),('" & TBAbrir!Nome_relatorio & "', 'Instutilizado')"
        Case "CQ_plano medicao_peca.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'Peças'), ('" & TBAbrir!Nome_relatorio & "', 'Instutilizado.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Encontrado')"
        Case "CQ_PPAP_FMEA.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_PPAP_Plano de controle.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Maquina')"
        Case "CQ_PPAP_Plano de controle_dimensional.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_relatorio_custo_retrabalho.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_relatorio_devolucao_clientes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_relatorio_inspecao_final.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_relatorio_inspecao_recebimento.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_relatorio_produzida_x_refugada.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_RNC.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_SA.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio'), ('" & TBAbrir!Nome_relatorio & "', 'Equipe')"
        Case "CQ_SD.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "CQ_RNC_lista.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "CQ_ultra_som.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Custos_centro_de_custo.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Custos_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'FaturamentoProd'), ('" & TBAbrir!Nome_relatorio & "', 'FaturamentoServ'), ('" & TBAbrir!Nome_relatorio & "', 'Fabricados'), ('" & TBAbrir!Nome_relatorio & "', 'Resultados Final'), ('" & TBAbrir!Nome_relatorio & "', 'Familia'), ('" & TBAbrir!Nome_relatorio & "', 'OS')"
        Case "Custos_previsto_realizado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "Custos_previsto_realizado_CC.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "Custos_previsto_realizado_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio'), ('" & TBAbrir!Nome_relatorio & "', 'Previsto'), ('" & TBAbrir!Nome_relatorio & "', 'Realizado')"
        Case "Custos_previsto_realizado_justificativa.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Custos_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Custos_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Custos_resumido_ordem.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Engenharia_conjuntos_estrutura.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Engenharia_estrutura.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Estrutura')"
        Case "Engenharia_familia.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Engenharia_produtos_servicos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Imagem')"
        Case "Engenharia_produtos_servicos_personalizado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Engenharia_normas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_fisico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_fisico_SemInventario.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_movimentacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_necessidade_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_necessidade_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_necessidade_resumido_vendas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_obsolescencia.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_ordemfaturamento_carteiraPC.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_ordemfaturamento_carteiraPI.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_ordemfaturamento_identificacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'Empenhos'), ('" & TBAbrir!Nome_relatorio & "', 'Empenhos_semOF')"
        Case "Estoque_ordemfaturamento_identificacao_personalizada.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'Empenhos'), ('" & TBAbrir!Nome_relatorio & "', 'Empenhos_semOF')"
        Case "Estoque_recebimento_pedido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Estoque_recebimento.rpt')"
        Case "Estoque_recebimento_plano inspecao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Familias.rpt')"
        Case "Estoque_recebimento_programacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Estoque_recebimento')"
        Case "Estoque_requisicao_materiais.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_saldo_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_saldo_diario.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_saldo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_semiacabado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Estoque_semiacabado_SPED.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Saldo.rpt')"
        Case "Eventos_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Eventos_detalhado_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Eventos_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Eventos_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Eventos_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Eventos_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Faturamento_carta correcao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Faturamento_impostos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Faturamento_minuta.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Produtos')"
        Case "Faturamento_nota fiscal_periodo_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_nota fiscal_periodo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_nota fiscal_produtos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'Duplicatas.rpt')"
        Case "Faturamento_nota fiscal_totais_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa.rpt')"
        Case "Faturamento_relacionamento.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio'), ('" & TBAbrir!Nome_relatorio & "', 'Relacionamento'), ('" & TBAbrir!Nome_relatorio & "', 'Relacionamento_NFcomplemento')"
        Case "Faturamento_relatorio_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_relatorio_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_relatorio_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_relatorio_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_relatorio_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa')"
        Case "Faturamento_romaneio.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Fiscal_CFOP.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Fiscal_Classificacao fiscal.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Fornecedores.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contatos')"
        Case "Fornecedores_lista_email.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contatos')"
        Case "Fornecedores_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Indice_atraso_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Indice_atraso_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Indice_atraso_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Indice_atraso_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Indice_atraso_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Instituicoes_cheque.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Instituicoes_extrato bancario.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Instituicoes_extrato bancario_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contas_pagas.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contas_pagas_transf.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contas_recebidas.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Contas_recebidas_transf.rpt')"
        Case "Instituicoes_extrato bancario_saldos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Instituicoes_movimentacao_financeira.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'CC'), ('" & TBAbrir!Nome_relatorio & "', 'CC_rec')"
        Case "Manutencao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Itens'), ('" & TBAbrir!Nome_relatorio & "', 'Check')"
        Case "Manutencao_cronograma.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'datas')"
        Case "Manutencao_relatorio_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Manutencao_relatorio_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Manutencao_relatorio_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Manutencao_relatorio_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Outros_SolicitacaoPCP.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_cargapostodetrabalho.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_cargapostodetrabalho_listaordem.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_carteira de pedidos_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Ordens')"
        Case "Pcp_carteira de pedidos_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'OFs')"
        Case "Pcp_eventosCB.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_indice_atraso_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_indice_atraso_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_indice_atraso_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_indice_atraso_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_indice_atraso_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_lista de ordens_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_lista de ordens_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Pedidos.rpt')"
        Case "PCP_nc_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "PCP_nc_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "PCP_nc_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "PCP_nc_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_ordem.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Pedidos'), ('" & TBAbrir!Nome_relatorio & "', 'Programas'), ('" & TBAbrir!Nome_relatorio & "', 'Ferramentas'), ('" & TBAbrir!Nome_relatorio & "', 'Plano de inspeção')"
        Case "Pcp_ordem_selecionadas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_ordem e rm.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Pedidos'), ('" & TBAbrir!Nome_relatorio & "', 'Materiais'), ('" & TBAbrir!Nome_relatorio & "', 'Programas'), ('" & TBAbrir!Nome_relatorio & "', 'Ferramentas'), ('" & TBAbrir!Nome_relatorio & "', 'Plano de inspeção')"
        Case "Pcp_ordem e rm_selecionadas.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Materiais')"
        Case "Pcp_ordem_apontamento manual.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Pedidos'), ('" & TBAbrir!Nome_relatorio & "', 'Programas'), ('" & TBAbrir!Nome_relatorio & "', 'Ferramentas'), ('" & TBAbrir!Nome_relatorio & "', 'Plano de inspeção')"
        Case "Pcp_plano da producao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_plano inspecao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Consignacao.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Familias.rpt')"
        Case "Pcp_plano inspecao_frequencia de medicao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Familias.rpt')"
        Case "Pcp_posto de trabalho_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Turnos')"
        Case "Pcp_posto de trabalho_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_programacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_material.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_resultados da ordem_material_NF.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Pcp_rm.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Pedidos'), ('" & TBAbrir!Nome_relatorio & "', 'Materiais.rpt')"
        Case "Pcp_situacao_producao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Fases')"
        Case "Processos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Ferramentas'), ('" & TBAbrir!Nome_relatorio & "', 'Programas')"
        Case "Processos_historico de fabricacao do item.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Produtividade_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Produtividade_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Produtividade_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Produtividade_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Produtividade_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "RH_funcionarios.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "RH_funcionarios_aniversariantes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "RH_funcionarios_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Cursos'), ('" & TBAbrir!Nome_relatorio & "', 'Aumentos'), ('" & TBAbrir!Nome_relatorio & "', 'Ferias'), ('" & TBAbrir!Nome_relatorio & "', 'Sindicato'), ('" & TBAbrir!Nome_relatorio & "', 'Atestados'), ('" & TBAbrir!Nome_relatorio & "', 'OBS')"
        Case "RH_relatorio_desoneracao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Suporte tecnico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Telemarketing_Visita.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Telemarketing_Visita tecnica.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Validacao_procedimentos.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Validacao_compras'), ('" & TBAbrir!Nome_relatorio & "', 'Validacao_ordem'), ('" & TBAbrir!Nome_relatorio & "', 'Data_insp_final'), ('" & TBAbrir!Nome_relatorio & "', 'Data_entrada_est'), ('" & TBAbrir!Nome_relatorio & "', 'Validacao_NF'), ('" & TBAbrir!Nome_relatorio & "', 'Data_expedicao')"
        Case "Vendas_alteracoes.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_analise.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'Engenharia') , ('" & TBAbrir!Nome_relatorio & "', 'Check-list_engenharia'), ('" & TBAbrir!Nome_relatorio & "', 'Processo'), ('" & TBAbrir!Nome_relatorio & "', 'Check-list_PCP'), ('" & TBAbrir!Nome_relatorio & "', 'Qualidade'), ('" & TBAbrir!Nome_relatorio & "', 'Check-list_qualidade'), ('" & TBAbrir!Nome_relatorio & "', 'Check-list_compras'), ('" & TBAbrir!Nome_relatorio & "', 'UltimoVlrVenda'), ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio')"
        Case "Vendas_comissao_comparativo_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_comissao_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_comissao_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_comissao_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_comissao_individual_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_follow up_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa'), ('" & TBAbrir!Nome_relatorio & "', 'Notafiscal')"
        Case "Vendas_follow up_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Empresa'), ('" & TBAbrir!Nome_relatorio & "', 'OFs')"
        Case "Vendas_controle detalhado de vendas_vi_dv.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_historico_resumido.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_historico_comparativo_resumido grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_historico_individual_detalhado.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_historico_individual_detalhado grafico.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt')"
        Case "Vendas_pedidointerno.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Listaprodutos.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Listaservicos.rpt')"
        Case "Vendas_pedidointerno_check list.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Lista de materiais')"
        Case "Vendas_programacao.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Recebimento'), ('" & TBAbrir!Nome_relatorio & "', 'Programacaomes'), ('" & TBAbrir!Nome_relatorio & "', 'Programacaoprevista')"
        Case "Vendas_proposta.rpt": Conexao.Execute "INSERT INTO Qualidade_revisao_relatorios_subreports (Nome_relatorio, Subreport) VALUES ('" & TBAbrir!Nome_relatorio & "', 'RevisaoRelatorio.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Listaprodutos.rpt'), ('" & TBAbrir!Nome_relatorio & "', 'Listaservicos.rpt')"
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Rel()
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
                If USMsgBox("Deseja realmente excluir este(s) relatório(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Qualidade_revisao_relatorios where Nome_relatorio = '" & .ListItems(InitFor).SubItems(1) & "'"
            Conexao.Execute "DELETE from Qualidade_revisao_relatorios_subreports where Nome_relatorio = '" & .ListItems(InitFor).SubItems(1) & "'"

            '==================================
            Modulo = "Qualidade/Histórico de revisão dos relatórios"
            Evento = "Excluir relatório"
            ID_documento = .ListItems(InitFor)
            Documento = "Nome do relatório: " & .ListItems(InitFor).SubItems(1)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) relatório(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Relatório(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Rel
    ProcCarregaLista (1)
    Frame3.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir_Revisao()
On Error GoTo tratar_erro
    
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) revisão(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Qualidade_revisao_relatorios where id = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Qualidade/Histórico de revisão dos relatórios"
            Evento = "Excluir revisão"
            ID_documento = .ListItems(InitFor)
            Documento = "Nome do relatório: " & .ListItems(InitFor).SubItems(1)
            Documento1 = "Nº revisão: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) revisão(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Revisão(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Revisao
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    ProcCarregaListaRevisao
    Frame2.Enabled = False
    Conexao.Execute "DELETE from Qualidade_revisao_relatorios_subreports where Nome_relatorio = '" & Txt_nome_relatorio & "'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir_Subreport()
On Error GoTo tratar_erro
    
If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) subreport(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Qualidade_revisao_relatorios_subreports where id = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Qualidade/Histórico de revisão dos relatórios"
            Evento = "Excluir subreport"
            ID_documento = .ListItems(InitFor)
            Documento = "Nome do relatório: " & Txt_nome_relatorio
            Documento1 = "Nome de subreport: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) subreport(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Subreport(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Subreport
    ProcCarregaListaSubreport
    Frame1.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar_Rel()
On Error GoTo tratar_erro

Acao = "salvar"
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txt_Caminho = "" Then
    NomeCampo = "o caminho do relatório"
    ProcVerificaAcao
    cmdImportar.SetFocus
    Exit Sub
End If
If Txt_nome_relatorio = "" Then
    NomeCampo = "o nome do relatório"
    ProcVerificaAcao
    Txt_nome_relatorio.SetFocus
    Exit Sub
End If
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios where id <> " & txtId & " and Nome_relatorio = '" & Txt_nome_relatorio & "' and Revisao = '" & Txt_revisao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    USMsgBox ("Este relatório já está cadastrado."), vbExclamation, "CAPRIND v5.0"
    TBHistProc.Close
    Exit Sub
End If

Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = True Then
    TBHistProc.AddNew
    TBHistProc!Revisao = 0
    TBHistProc!Data = Date
    TBHistProc!Responsavel = pubUsuario
End If
TBHistProc!caminho = txt_Caminho
TBHistProc!Nome_relatorio = Txt_nome_relatorio
If Chk_personalizado.Value = 1 Then TBHistProc!Personalizado = True Else TBHistProc!Personalizado = False
If chkResponsavel.Value = 1 Then TBHistProc!Responsavel_rel = True Else TBHistProc!Responsavel_rel = False
TBHistProc.Update
txtId = TBHistProc!ID
Caption = "Qualidade - Histórico de revisão dos relatórios - (Nome do relatório :  " & IIf(IsNull(TBHistProc!Nome_relatorio), "", TBHistProc!Nome_relatorio) & " - Revisão : " & IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao) & ")"
TBHistProc.Close
If Novo_Revisao_Relatorios = True Then
    USMsgBox ("Novo relatório cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo relatório"
    Sql_Revisao_Relatorios_Localizar = "Select Nome_relatorio from Qualidade_revisao_relatorios where ID = " & txtId & " group by Nome_relatorio"
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar relatório"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Histórico de revisão dos relatórios"
ID_documento = txtId
Documento = "Nome do relatório: " & Txt_nome_relatorio
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Revisao_Relatorios = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar_Revisao()
On Error GoTo tratar_erro

Acao = "salvar"
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Txt_revisao = "" Then
    NomeCampo = "a revisão"
    ProcVerificaAcao
    Txt_revisao.SetFocus
    Exit Sub
End If
If Txt_responsavel = "" Then
    NomeCampo = "o responsável"
    ProcVerificaAcao
    Txt_responsavel.SetFocus
    Exit Sub
End If
If Txt_historico = "" Then
    NomeCampo = "o histórico"
    ProcVerificaAcao
    Txt_historico.SetFocus
    Exit Sub
End If

Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios where id = " & txtId1, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = True Then TBHistProc.AddNew
TBHistProc!caminho = txt_Caminho
TBHistProc!Nome_relatorio = Txt_nome_relatorio
TBHistProc!Revisao = Txt_revisao
TBHistProc!Data = Txt_data_revisao
TBHistProc!Responsavel = Txt_responsavel
TBHistProc!Historico = Txt_historico
TBHistProc.Update
txtId1 = TBHistProc!ID
TBHistProc.Close
ProcCarregaListaRevisao
If Novo_Revisao_Relatorios1 = True Then
    USMsgBox ("Nova revisão cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova revisão"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar revisão"
    If CodigoLista1 <> 0 And Lista1.ListItems.Count <> 0 Then
        Lista1.SelectedItem = Lista1.ListItems(CodigoLista1)
        Lista1.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Histórico de revisão dos relatórios"
ID_documento = txtId1
Documento = "Nome do relatório: " & Txt_nome_relatorio
Documento1 = "Nº revisão: " & Txt_revisao
ProcGravaEvento
'==================================
Novo_Revisao_Relatorios1 = False
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar_Subreport()
On Error GoTo tratar_erro

Acao = "salvar"
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Txt_nome_subreport = "" Then
    NomeCampo = "o nome do subreport"
    ProcVerificaAcao
    Txt_nome_subreport.SetFocus
    Exit Sub
End If

Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Qualidade_revisao_relatorios_subreports where id = " & txtID2, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = True Then TBHistProc.AddNew
TBHistProc!Nome_relatorio = Txt_nome_relatorio
TBHistProc!SubReport = Txt_nome_subreport
TBHistProc.Update
txtID2 = TBHistProc!ID
TBHistProc.Close
ProcCarregaListaSubreport
If Novo_Revisao_Relatorios2 = True Then
    USMsgBox ("Novo subreport cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo subreport"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar subreport"
    If CodigoLista2 <> 0 And Lista2.ListItems.Count <> 0 Then
        Lista2.SelectedItem = Lista2.ListItems(CodigoLista2)
        Lista2.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Histórico de revisão dos relatórios"
ID_documento = txtID2
Documento = "Nome do relatório: " & Txt_nome_relatorio
Documento1 = "Nome do subreport: " & Txt_nome_subreport
ProcGravaEvento
'==================================
Novo_Revisao_Relatorios2 = False
ProcCarregaListaSubreport

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
