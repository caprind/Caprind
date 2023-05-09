VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_trocaduplicata 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Desconto de duplicata"
   ClientHeight    =   10035
   ClientLeft      =   195
   ClientTop       =   465
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
   Begin MSComctlLib.ListView ListaPrincipal 
      Height          =   5655
      Left            =   75
      TabIndex        =   74
      Top             =   3440
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   9975
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
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Borderô"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Banco recebedor"
         Object.Width           =   7853
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Local do desconto"
         Object.Width           =   7853
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Vlr. resgatado"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.ComboBox cmbbanco 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Banco recebedor."
      Top             =   2255
      Width           =   6435
   End
   Begin VB.ComboBox txtlocaltroca 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9280
      Sorted          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Local do desconto."
      Top             =   2255
      Width           =   5805
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
      ItemData        =   "frm_trocaduplicata.frx":0000
      Left            =   240
      List            =   "frm_trocaduplicata.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1690
      Width           =   6435
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   28
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
   Begin MSComctlLib.ListView Lista 
      Height          =   6675
      Left            =   75
      TabIndex        =   17
      Top             =   2190
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   11774
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. venc."
         Object.Width           =   2117
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
         Text            =   "Nº docto."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   7770
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3881
      EndProperty
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
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
      TabCaption(0)   =   "Borderôs"
      TabPicture(0)   =   "frm_trocaduplicata.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Duplicatas"
      TabPicture(1)   =   "frm_trocaduplicata.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtidconta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FrameDuplicata"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   63
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
            TabIndex        =   65
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
            TabIndex        =   64
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   66
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_trocaduplicata.frx":003C
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
            TabIndex        =   67
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_trocaduplicata.frx":37E0
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
            TabIndex        =   68
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
            TabIndex        =   69
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_trocaduplicata.frx":72E9
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
            TabIndex        =   70
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_trocaduplicata.frx":B3D8
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
            Left            =   3510
            TabIndex        =   75
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   73
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
            TabIndex        =   72
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   2190
            TabIndex        =   71
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2115
         Left            =   75
         TabIndex        =   51
         Top             =   1305
         Width           =   15195
         Begin VB.TextBox txtSaldoAtual 
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
            Left            =   6630
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Saldo anterior."
            Top             =   945
            Width           =   1275
         End
         Begin VB.TextBox txtSaldo 
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
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Saldo."
            Top             =   945
            Width           =   1275
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
            Left            =   9215
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   390
            Width           =   5795
         End
         Begin VB.TextBox txtBordero 
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
            Left            =   6630
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número do borderô."
            Top             =   390
            Width           =   1275
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
            Height          =   465
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            ToolTipText     =   "Observações."
            Top             =   1515
            Width           =   14805
         End
         Begin MSComCtl2.DTPicker txtData 
            Height          =   315
            Left            =   7920
            TabIndex        =   2
            ToolTipText     =   "Data de emissão."
            Top             =   390
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
            Format          =   198639617
            CurrentDate     =   39057
         End
         Begin VB.Label Label27 
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
            Left            =   2880
            TabIndex        =   60
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo anterior"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6765
            TabIndex        =   59
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8362
            TabIndex        =   58
            Top             =   750
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11630
            TabIndex        =   57
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8385
            TabIndex        =   56
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Borderô"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6982
            TabIndex        =   55
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observação"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7147
            TabIndex        =   54
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco recebedor*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2670
            TabIndex        =   53
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local do desconto*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11490
            TabIndex        =   52
            Top             =   750
            Width           =   1380
         End
      End
      Begin VB.Frame FrameDuplicata 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   -74925
         TabIndex        =   42
         Top             =   1305
         Width           =   15195
         Begin VB.TextBox txtpmedio 
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
            Left            =   150
            TabIndex        =   9
            ToolTipText     =   "Prazo médio."
            Top             =   420
            Width           =   1425
         End
         Begin VB.TextBox txtenviado 
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
            Left            =   12000
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Valor enviado."
            Top             =   420
            Width           =   2985
         End
         Begin VB.TextBox txtretido 
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
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Valor retido."
            Top             =   420
            Width           =   1875
         End
         Begin VB.TextBox txtpis 
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
            Left            =   3930
            TabIndex        =   11
            ToolTipText     =   "Valor do PIS/Cofins."
            Top             =   420
            Width           =   1305
         End
         Begin VB.TextBox txtvalortitulo 
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
            Left            =   1590
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Valor da duplicata."
            Top             =   420
            Width           =   2325
         End
         Begin VB.TextBox txtCofins 
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
            Left            =   5250
            TabIndex        =   12
            ToolTipText     =   "Valor do PIS/Cofins."
            Top             =   420
            Width           =   1305
         End
         Begin VB.TextBox txtVlr_pis 
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
            Left            =   6570
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Valor retido do pis."
            Top             =   420
            Width           =   1755
         End
         Begin VB.TextBox txtVlr_cofins 
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
            Left            =   8340
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Valor retido do cofins."
            Top             =   420
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo médio (dias)*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   157
            TabIndex        =   50
            Top             =   210
            Width           =   1410
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor retido"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10635
            TabIndex        =   49
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor enviado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13005
            TabIndex        =   48
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PIS %*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4312
            TabIndex        =   47
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor da duplicata"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2115
            TabIndex        =   46
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COFINS %*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5467
            TabIndex        =   45
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor retido COFINS"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   8497
            TabIndex        =   44
            Top             =   210
            Width           =   1440
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor retido PIS"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6892
            TabIndex        =   43
            Top             =   210
            Width           =   1110
         End
      End
      Begin VB.TextBox txtidconta 
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
         Left            =   -74100
         TabIndex        =   41
         Text            =   "0"
         ToolTipText     =   "Valor do PIS/Cofins."
         Top             =   3150
         Visible         =   0   'False
         Width           =   645
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
         Height          =   825
         Left            =   -74925
         TabIndex        =   30
         Top             =   8880
         Width           =   15195
         Begin VB.TextBox Txt_prazo 
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
            Left            =   12045
            TabIndex        =   25
            ToolTipText     =   "Prazo."
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox Txt_taxa_mes 
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
            Left            =   13215
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Taxa de juros ao mês."
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtvlrtotalenviado 
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
            Left            =   2460
            TabIndex        =   19
            ToolTipText     =   "Valor total enviado."
            Top             =   360
            Width           =   2265
         End
         Begin VB.TextBox txtvlrtotalretido 
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
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Valor total retido."
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txttaxatotal 
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
            Left            =   11295
            TabIndex        =   24
            ToolTipText     =   "Taxa de juros no período."
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtvlrtotaltitulo 
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
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de duplicatas."
            Top             =   360
            Width           =   2265
         End
         Begin VB.TextBox txtvlrtotalresgatado 
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
            Left            =   7785
            TabIndex        =   22
            ToolTipText     =   "Valor total resgatado."
            Top             =   360
            Width           =   1875
         End
         Begin VB.CommandButton cmdImpostos 
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
            Left            =   7110
            Picture         =   "frm_trocaduplicata.frx":EC64
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cadastrar impostos."
            Top             =   360
            Width           =   315
         End
         Begin MSMask.MaskEdBox Txt_data_operacao 
            Height          =   315
            Left            =   9675
            TabIndex        =   23
            ToolTipText     =   "Data da operação."
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin DrawSuite2022.USButton CmdDescDuplicata 
            Height          =   480
            Left            =   14250
            TabIndex        =   27
            ToolTipText     =   "Descontar duplicata."
            Top             =   210
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   847
            Caption         =   "Descontar"
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
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo (dias)"
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
            Left            =   12112
            TabIndex        =   40
            Top             =   180
            Width           =   1020
         End
         Begin VB.Image Imgcalendario 
            Height          =   360
            Left            =   10950
            MouseIcon       =   "frm_trocaduplicata.frx":ED46
            Picture         =   "frm_trocaduplicata.frx":F050
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Taxa mês"
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
            Left            =   13290
            TabIndex        =   39
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. operação"
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
            Left            =   9772
            TabIndex        =   38
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Left            =   4845
            TabIndex        =   37
            Top             =   420
            Width           =   75
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
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
            Left            =   7545
            TabIndex        =   36
            Top             =   420
            Width           =   135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Taxa"
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
            Left            =   11475
            TabIndex        =   35
            Top             =   180
            Width           =   420
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. total retido"
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
            Left            =   5430
            TabIndex        =   34
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. total resgatado"
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
            Left            =   7905
            TabIndex        =   33
            Top             =   180
            Width           =   1635
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. total enviado"
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
            Left            =   2872
            TabIndex        =   32
            Top             =   180
            Width           =   1440
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. total duplicatas"
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
            TabIndex        =   31
            Top             =   180
            Width           =   1635
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   61
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   12
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
         ButtonKey8      =   "9"
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
         ButtonLeft9     =   358
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
         ButtonLeft10    =   362
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
         ButtonLeft11    =   400
         ButtonTop11     =   2
         ButtonWidth11   =   26
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   5
         ButtonLeft12    =   428
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   13680
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frm_trocaduplicata.frx":F4D3
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   62
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   33
         ButtonHeight1   =   21
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
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
         ButtonLeft3     =   77
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   51
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   171
         ButtonTop5      =   2
         ButtonWidth5    =   47
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   220
         ButtonTop6      =   2
         ButtonWidth6    =   46
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
         ButtonLeft7     =   268
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   272
         ButtonTop8      =   2
         ButtonWidth8    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   310
         ButtonTop9      =   2
         ButtonWidth9    =   26
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
         ButtonLeft10    =   338
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13680
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frm_trocaduplicata.frx":1614C
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frm_trocaduplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VlrRetido                      As Double 'OK
Dim PIS                            As Double 'OK
Dim Enviado                        As Double 'OK
Public Novo_Desconto               As Boolean 'OK
Public TemImposto                  As Boolean 'OK
Public StrSql_Desconto_Duplicata   As String 'OK
Dim TBLISTA_Desconto_Duplicata     As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=dXhVsaexDhk&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=21&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
ProcCarregaCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbbanco_Click()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from tbl_instituicoes where txt_descricao = '" & cmbBanco & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from troca_titulo where id = " & IIf(txtBordero = "", 0, txtBordero), Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        txtSaldoAtual = Format(TBFIltro!Saldo - IIf(IsNull(TBFI!Vlrtotalresgatado), 0, TBFI!Vlrtotalresgatado), "###,##0.00")
    Else
        txtSaldoAtual = Format(TBFIltro!Saldo, "###,##0.00")
    End If
    TBFI.Close
    ProcAtualizaSaldo
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtBordero = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from troca_titulo order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtBordero)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtBordero = TBLISTA!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from troca_titulo where ID = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaValores
        ProcLimpaCamposTotais
        ProcCarregaDados
        ProcCarregaLista
        ProcCarregaTotais
        Frame4.Enabled = True
        Frame1.Enabled = True
        cmbBanco.Enabled = True
        txtlocaltroca.Enabled = True
    Else
        USMsgBox ("Fim dos cadastros de borderôs."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Desconto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdDescDuplicata_Click()
On Error GoTo tratar_erro

Acao = "descontar a(s) duplicata(s)"
If txtBordero = "" Then
    NomeCampo = "o borderô"
    ProcVerificaAcao
    Exit Sub
End If
If Lista.ListItems.Count = 0 Then
    USMsgBox ("Adicione uma duplicata na lista antes de descontar a(s) duplicata(s)"), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = True
ProcVerificaStatusDuplicatas
If Permitido = False Then Exit Sub
If txtvlrtotalresgatado = "" Then
    NomeCampo = "o valor total resgatado"
    ProcVerificaAcao
    txtvlrtotalresgatado.SetFocus
    Exit Sub
End If
If txtvlrtotalresgatado <> "0" And txtvlrtotalresgatado <> "0,00" Then
    If Txt_data_operacao = "__/__/____" Then
        NomeCampo = "a data da operação"
        ProcVerificaAcao
        Txt_data_operacao.SetFocus
        Exit Sub
    End If
End If
If USMsgBox("Deseja realmente descontar essas(s) duplicata(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from troca_titulo where id = " & IIf(txtBordero = "", 0, txtBordero), Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    ProcSubtraiVlrResgatadoAnterior
    ProcAtualizaSaldoBancario
    Evento = "Descontar duplicata"
End If
TBContas!Data = txtData.Value
If txtResponsavel <> "" Then TBContas!Responsavel = txtResponsavel.Text Else TBContas!Responsavel = pubUsuario
TBContas!local_troca = txtlocaltroca.Text
TBContas!banco_recebedor = cmbBanco.Text
TBContas!Obs = txtObs.Text
TBContas!Vlrtotaltitulo = IIf(txtvlrtotaltitulo.Text = "", 0, txtvlrtotaltitulo.Text)
TBContas!Vlrtotalenviado = IIf(txtvlrtotalenviado.Text = "", 0, txtvlrtotalenviado.Text)
TBContas!vlrtotalretido = IIf(txtvlrtotalretido.Text = "", 0, txtvlrtotalretido.Text)
TBContas!Vlrtotalresgatado = IIf(txtvlrtotalresgatado.Text = "", 0, txtvlrtotalresgatado.Text)
TBContas!Data_operacao = IIf(Txt_data_operacao = "__/__/____", Null, Txt_data_operacao)
TBContas!Taxatotal = IIf(txttaxatotal.Text = "", 0, txttaxatotal.Text)
TBContas!Pmedio = IIf(Txt_prazo = "", 0, Txt_prazo)
If txtvlrtotalresgatado <> "" And txtvlrtotalresgatado <> "0" And txtvlrtotalresgatado <> "0,00" Then
    SaldoAtual = txtSaldoAtual
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select tbl_Fluxo_de_caixa.* from tbl_contas_receber INNER JOIN tbl_Fluxo_de_caixa on tbl_contas_receber.IDFluxo = tbl_Fluxo_de_caixa.IDFluxo where tbl_contas_receber.IDtrocatitulo = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        Do While TBFluxo.EOF = False
            TBFluxo!Operacao = "Crédito"
            TBFluxo!Data = Txt_data_operacao
            TBFluxo!Instituicao = cmbBanco
            TBFluxo!Hora = Format(Now, "hh:mm:ss")
            TBFluxo!status = "S"
            TBFluxo!Cheque = txtBordero
            TBFluxo!Bloqueado = True
            TBFluxo.Update
            TBFluxo.MoveNext
        Loop
    End If
    TBFluxo.Close
    
    'Fluxo de Caixa
    'Cria registro com o valor total da operação
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = True Then
        TBFluxo.AddNew
        TBFluxo!Descricao = "Desconto de duplicata borderô n. " & txtBordero
        TBFluxo!Obs = "Desconto de duplicata borderô n. " & txtBordero
    End If
    TBFluxo!Operacao = "Crédito"
    TBFluxo!Data = Txt_data_operacao
    TBFluxo!valor = txtvlrtotalresgatado
    TBFluxo!Instituicao = cmbBanco
    TBFluxo!status = "S"
    TBFluxo!Hora = Format(Now, "hh:mm:ss")
    TBFluxo!Cheque = txtBordero
    TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBFluxo!Bloqueado = False
    TBFluxo.Update
    TBContas!IDFluxo = TBFluxo!IDFluxo
    TBFluxo.Close
Else
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select tbl_Fluxo_de_caixa.*, tbl_contas_receber.Vencimento , tbl_contas_receber.Valor as Valor1 from tbl_contas_receber INNER JOIN tbl_Fluxo_de_caixa on tbl_contas_receber.IDFluxo = tbl_Fluxo_de_caixa.IDFluxo where tbl_contas_receber.IDtrocatitulo = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        Do While TBFluxo.EOF = False
            TBFluxo!Operacao = "À Creditar"
            TBFluxo!Data = TBFluxo!Vencimento
            TBFluxo!valor = TBFluxo!Valor1
            TBFluxo!Instituicao = Null
            TBFluxo!Hora = Null
            TBFluxo!status = "N"
            TBFluxo!Cheque = 0
            TBFluxo!Bloqueado = False
            TBFluxo.Update
            TBFluxo.MoveNext
        Loop
    End If
    TBFluxo.Close
    
    'Fluxo de Caixa
    'Exclui registro com o valor total da operação
    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
End If
TBContas.Update
TBContas.Close
USMsgBox ("Duplicata(s) descontada(s) com sucesso."), vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImpostos_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Or txtIDConta = "" Then Exit Sub
frm_trocaduplicata_Valores.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frm_trocaduplicata_filtro.Show 1

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
ProcLimpaValores
ProcLimpaCamposTotais
Lista.ListItems.Clear
Novo_Desconto = True
TemImposto = False
Frame4.Enabled = True
Frame1.Enabled = True
cmbBanco.Enabled = True
txtlocaltroca.Enabled = True
txtvlrtotalresgatado.Locked = False
txtvlrtotalresgatado.TabStop = True
cmbBanco.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtBordero = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from troca_titulo order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtBordero)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtBordero = TBLISTA!ID
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from troca_titulo where ID = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaValores
        ProcLimpaCamposTotais
        ProcCarregaDados
        ProcCarregaLista
        ProcCarregaTotais
        Frame4.Enabled = True
        cmbBanco.Enabled = True
        txtlocaltroca.Enabled = True
        Frame1.Enabled = True
    Else
        USMsgBox ("Fim dos cadastros de borderôs."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Desconto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarLista()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtIDConta = 0 Then
    ProcVerificaSalvar
    Exit Sub
End If
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select IDIntconta from tbl_contas_receber where IDIntconta = " & txtIDConta & " and (Status = 'DUPLICATA DESCONTADA LIQUIDADA' or Status = 'DUPLICATA DESCONTADA RECOMPRADA')", Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    USMsgBox ("Não é permitido salvar esta duplicata pois a mesma está liquídada ou recomprada."), vbInformation, "CAPRIND v5.0"
    TBContas.Close
    Exit Sub
End If
TBContas.Close
Acao = "salvar"
If txtpmedio = "" Then
    NomeCampo = "o prazo médio"
    ProcVerificaAcao
    txtpmedio.SetFocus
    Exit Sub
End If
If txtPIS = "" Then
    NomeCampo = "o valor do pis"
    ProcVerificaAcao
    txtPIS.SetFocus
    Exit Sub
End If
If txtCofins = "" Then
    NomeCampo = "o valor do cofins"
    ProcVerificaAcao
    txtCofins.SetFocus
    Exit Sub
End If

Valorenviado = IIf(txtvlrtotalenviado = "", 0, txtvlrtotalenviado)
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from troca_titulo_valores where N_conta = " & txtIDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!IDduplicata = txtBordero
TBGravar!n_conta = txtIDConta
TBGravar!valor_enviado = txtenviado.Text
TBGravar!valor_retido = txtretido.Text
TBGravar!valor_pis = txtPIS.Text
TBGravar!valor_cofins = txtCofins
TBGravar!Prazo = txtpmedio
TBGravar.Update
TBGravar.Close
ProcGravarTotais
USMsgBox ("Duplicata cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaLimiteUtil()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & txtlocaltroca & "' and tbl_contas_receber.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    NovoValor = Replace(valor, ",", ".")
    Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & txtlocaltroca & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
txtBordero = TBAbrir!ID
If IsNull(TBAbrir!Data) = False Then txtData.Value = TBAbrir!Data
txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
If IsNull(TBAbrir!banco_recebedor) = False And TBAbrir!banco_recebedor <> "" Then cmbBanco.Text = TBAbrir!banco_recebedor
txtObs.Text = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
txtlocaltroca.Text = IIf(IsNull(TBAbrir!local_troca), "", TBAbrir!local_troca)
'Verifica se tem imposto cadastrado para este bordero

With txtvlrtotalresgatado
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select ID from troca_titulo_ValoresImpostos where ID_duplicata = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        TemImposto = True
        .Locked = True
        .TabStop = False
    Else
        TemImposto = False
        .Locked = False
        .TabStop = True
    End If
    TBFI.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarTotais()
On Error GoTo tratar_erro

Enviado = 0
VlrRetido = 0
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Sum(valor_enviado) as Enviado, Sum(valor_retido) as Vlrretido from troca_titulo_valores where IDduplicata = " & txtBordero.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Enviado = IIf(IsNull(TBFI!Enviado), 0, TBFI!Enviado)
    VlrRetido = IIf(IsNull(TBFI!VlrRetido), 0, TBFI!VlrRetido)
End If
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Sum(valor) as Vlrretido from troca_titulo_valoresImpostos where id_duplicata = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    VlrRetido = VlrRetido + IIf(IsNull(TBFI!VlrRetido), 0, TBFI!VlrRetido)
End If
TBFI.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from troca_titulo where id = " & txtBordero.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Vlrtotaltitulo = txtvlrtotaltitulo
TBGravar!Vlrtotalenviado = Enviado
TBGravar!vlrtotalretido = VlrRetido
TBGravar!Vlrtotalresgatado = IIf(txtvlrtotalresgatado = "", 0, txtvlrtotalresgatado)
TBGravar!Taxatotal = IIf(txttaxatotal = "", 0, txttaxatotal)
TBGravar.Update
TBGravar.Close

txtvlrtotalenviado.Text = Format(Enviado, "###,##0.00")
txtvlrtotalretido.Text = Format(VlrRetido, "###,##0.00")
Enviado = 0
VlrRetido = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaTotais()
On Error GoTo tratar_erro

txtvlrtotaltitulo.Text = IIf(IsNull(TBAbrir!Vlrtotaltitulo), "", Format(TBAbrir!Vlrtotaltitulo, "###,##0.00"))
txtvlrtotalenviado.Text = IIf(IsNull(TBAbrir!Vlrtotalenviado), "", Format(TBAbrir!Vlrtotalenviado, "###,##0.00"))
txtvlrtotalretido.Text = IIf(IsNull(TBAbrir!vlrtotalretido), "", Format(TBAbrir!vlrtotalretido, "###,##0.00"))
txtvlrtotalresgatado.Text = IIf(IsNull(TBAbrir!Vlrtotalresgatado), "", Format(TBAbrir!Vlrtotalresgatado, "###,##0.00"))
Txt_data_operacao = IIf(IsNull(TBAbrir!Data_operacao), "__/__/____", Format(TBAbrir!Data_operacao, "dd/mm/yyyy"))
txttaxatotal.Text = IIf(IsNull(TBAbrir!Taxatotal), "", Format(TBAbrir!Taxatotal, "###,##0.00"))
Txt_prazo = IIf(IsNull(TBAbrir!Pmedio), "", TBAbrir!Pmedio)
Total = 0

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
If Frame4.Enabled = False Then
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
If cmbBanco.Text = "" Then
    NomeCampo = "o banco recebedor"
    ProcVerificaAcao
    cmbBanco.SetFocus
    Exit Sub
End If
If txtlocaltroca.Text = "" Then
    NomeCampo = "o local da troca"
    ProcVerificaAcao
    txtlocaltroca.SetFocus
    Exit Sub
End If
If Lista.ListItems.Count <> 0 Then
    Permitido = True
    ProcVerificaStatusDuplicatas
    If Permitido = False Then Exit Sub
End If

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from troca_titulo where id = " & IIf(txtBordero = "", 0, txtBordero), Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    ProcSubtraiVlrResgatadoAnterior
    ProcAtualizaSaldoBancario
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
Else
    TBContas.AddNew
    ProcAtualizaSaldoBancario
    USMsgBox ("Novo borderô cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
End If
TBContas!Data = txtData.Value
If txtResponsavel <> "" Then TBContas!Responsavel = txtResponsavel.Text Else TBContas!Responsavel = pubUsuario
TBContas!local_troca = txtlocaltroca.Text
TBContas!banco_recebedor = cmbBanco.Text
TBContas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBContas!Obs = txtObs.Text
TBContas.Update
txtBordero = TBContas!ID
TBContas.Close
 
'==================================
Modulo = "Financeiro/Contas à receber/Desconto de duplicata"
ID_documento = txtBordero
Documento = "Borderô: " & txtBordero
Documento1 = ""
ProcGravaEvento
'==================================
If Novo_Desconto = True Then
    StrSql_Desconto_Duplicata = "Select * from troca_titulo where ID = " & txtBordero
    procCarregalistaPrincipal 1
Else
    procCarregalistaPrincipal (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And ListaPrincipal.ListItems.Count <> 0 Then
        ListaPrincipal.SelectedItem = ListaPrincipal.ListItems(CodigoLista)
        ListaPrincipal.SetFocus
    End If
End If
Novo_Desconto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSubtraiVlrResgatadoAnterior()
On Error GoTo tratar_erro

Set TBReceber = CreateObject("adodb.recordset")
TBReceber.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & cmbBanco.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBReceber.EOF = False Then
    TBReceber!Saldo = TBReceber!Saldo - IIf(IsNull(TBContas!Vlrtotalresgatado), 0, TBContas!Vlrtotalresgatado)
    TBReceber.Update
End If
TBReceber.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldoBancario()
On Error GoTo tratar_erro

Set TBReceber = CreateObject("adodb.recordset")
TBReceber.Open "Select * from tbl_Instituicoes where txt_Descricao = '" & cmbBanco.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBReceber.EOF = False Then
    TBReceber!Saldo = TBReceber!Saldo + IIf(txtvlrtotalresgatado.Text = "", 0, txtvlrtotalresgatado.Text)
    TBReceber.Update
End If
TBReceber.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Desconto_Duplicata.AbsolutePage <> 2 Then
    If TBLISTA_Desconto_Duplicata.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Desconto_Duplicata.PageCount - 1)
    Else
        TBLISTA_Desconto_Duplicata.AbsolutePage = TBLISTA_Desconto_Duplicata.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Desconto_Duplicata.AbsolutePage)
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
    TBLISTA_Desconto_Duplicata.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Desconto_Duplicata.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Desconto_Duplicata.AbsolutePage = 1
ProcExibePagina (TBLISTA_Desconto_Duplicata.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Desconto_Duplicata.AbsolutePage <> -3 Then
    If TBLISTA_Desconto_Duplicata.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Desconto_Duplicata.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Desconto_Duplicata.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Desconto_Duplicata.AbsolutePage = TBLISTA_Desconto_Duplicata.PageCount
ProcExibePagina (TBLISTA_Desconto_Duplicata.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 12, True
ProcCarregaToolBar2 Me, 15192, 10, True

Formulario = "Financeiro/Desconto de duplicata"
Direitos
ProcLimpaVariaveisPrincipais
txtData.Value = Date
ProcCarregaComboEmpresa Cmb_empresa, False
SSTab1.Tab = 0

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombo()
On Error GoTo tratar_erro

ProcCarregaComboBancoFinanceiro cmbBanco, "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
ProcCarregaComboBancoFinanceiro txtlocaltroca, "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Desconto de duplicata"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362D" Then frm_trocaduplicata_atualizar.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frm_trocaduplicata_atualizar
        If .Chk1.Value = 1 Then
            'Atualizar dados das duplicatas descontadas
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from troca_titulo order by id", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                TBContas.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBContas.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBContas.MoveFirst
                Do While TBContas.EOF = False
                    Contador1 = 0
                    PrazoMedio = 0
                    Valorenviado = 0
                    Set TBReceber = CreateObject("adodb.recordset")
                    TBReceber.Open "Select * from troca_titulo_valores where IDduplicata = " & TBContas!ID, Conexao, adOpenKeyset, adLockOptimistic
                    If TBReceber.EOF = False Then
                        Do While TBReceber.EOF = False
                            'Atualiza prazo medio da duplicata
                            TBReceber!Prazo = 0
                            
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from tbl_contas_receber where IDIntconta = " & TBReceber!n_conta, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                ValorTotal = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                                If IsNull(TBAbrir!Vencimento) = False And TBAbrir!Vencimento >= TBContas!Data Then Data = TBAbrir!Vencimento - TBContas!Data Else Data = 0
                                ElapsedTime (Data)
                                TBReceber!Prazo = D
                            End If
                            TBAbrir.Close
                            TBReceber.Update
                            
                            Contador1 = Contador1 + 1
                            PrazoMedio = PrazoMedio + (IIf(IsNull(TBReceber!valor_enviado), 0, TBReceber!valor_enviado) * IIf(IsNull(TBReceber!Prazo), 0, TBReceber!Prazo))
                            Valorenviado = Valorenviado + IIf(IsNull(TBReceber!valor_enviado), 0, TBReceber!valor_enviado)
                            TBReceber.MoveNext
                        Loop
                    End If
                    TBReceber.Close
                    TBContas!NDuplicata = Contador1
                    If PrazoMedio <> 0 And Valorenviado <> 0 Then PrazoMedio = PrazoMedio / Valorenviado
                    TBContas!Pmedio = PrazoMedio
                    If IsNull(TBContas!Vlrtotalrecompra) = True Or TBContas!Vlrtotalrecompra = "" Then TBContas!Vlrtotalrecompra = 0
                    If TBContas!Vlrtotalresgatado <> 0 And IsNull(TBContas!Data_operacao) = True Or TBContas!Vlrtotalresgatado <> 0 And TBContas!Data_operacao = "" Then TBContas!Data_operacao = TBContas!Data
                    TBContas.Update
                    TBContas.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBContas.Close
        End If
        
        If .Chk2.Value = 1 Then
            'Atualiza data de recebimento dos titulos descontados liquidados
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber where status = 'DUPLICATA DESCONTADA LIQUIDADA' order by Vencimento", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TBAbrir.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBAbrir.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBAbrir.MoveFirst
                Do While TBAbrir.EOF = False
                    TBAbrir!Data_pagamento = TBAbrir!Vencimento
                    TBAbrir.Update
                    TBAbrir.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBAbrir.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Atualiza ID do fluxo de caixa nos borderôs
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from troca_titulo where Data_operacao <> 'Null' order by id", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                TBContas.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBContas.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBContas.MoveFirst
                Do While TBContas.EOF = False
                    Set TBFluxo = CreateObject("adodb.recordset")
                    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Cheque = '" & TBContas!ID & "' and Instituicao = '" & TBContas!banco_recebedor & "' and Data = '" & Format(TBContas!Data_operacao, "dd/mm/YYYY") & "' and Valor = " & NovoValor, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFluxo.EOF = False Then
                        TBContas!IDFluxo = TBFluxo!IDFluxo
                        TBContas.Update
                    End If
                    TBFluxo.Close
                    TBContas.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBContas.Close
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Financeiro/Contas à receber/Desconto de duplicata"
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
SolicitacaoAcao = False
Troca_Duplicata = True
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
With ListaPrincipal
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) borderô(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from troca_titulo where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                IDFluxo = IIf(IsNull(TBAbrir!IDFluxo), 0, TBAbrir!IDFluxo)
                
                'Verifica se o bordero foi descontado para corrigir o saldo do banco
                Familiatext = "Desconto de duplicata borderô n. " & txtBordero
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select IDFluxo from tbl_Fluxo_de_caixa where Descricao = '" & Familiatext & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    Set TBReceber = CreateObject("adodb.recordset")
                    TBReceber.Open "Select Saldo from tbl_Instituicoes where txt_Descricao = '" & IIf(IsNull(TBAbrir!banco_recebedor), "", TBAbrir!banco_recebedor) & "' and ID_empresa = " & IIf(IsNull(TBAbrir!ID_empresa), 0, TBAbrir!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
                    If TBReceber.EOF = False Then
                        TBReceber!Saldo = TBReceber!Saldo - IIf(IsNull(TBAbrir!Vlrtotalresgatado), 0, TBAbrir!Vlrtotalresgatado)
                        TBReceber.Update
                    End If
                    TBReceber.Close
                End If
                TBFluxo.Close
            End If
            TBAbrir.Close
            Conexao.Execute "DELETE from troca_titulo where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from troca_titulo_valores where IDduplicata = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from troca_titulo_valoresImpostos where id_duplicata = " & .ListItems(InitFor)
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber where IDtrocatitulo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    'Fluxo de Caixa
                    Set TBFluxo = CreateObject("adodb.recordset")
                    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBAbrir!IDFluxo), 0, TBAbrir!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFluxo.EOF = False Then
                        TBFluxo!Operacao = "À Creditar"
                        TBFluxo!Data = TBAbrir!Vencimento
                        TBFluxo!valor = TBAbrir!valor
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!status = "N"
                        TBFluxo!Cheque = 0
                        TBFluxo!Bloqueado = False
                        TBFluxo.Update
                    End If
                    TBFluxo.Close
                    
                    TBAbrir!IDtrocatitulo = 0
                    'Verif. se a conta esta recebida parcial
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from tbl_contas_receber where tituloref = '" & TBAbrir!IDintconta & "' and idintconta <> " & TBAbrir!IDintconta & " and logsit = 'S'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBAbrir!status = "TÍTULO RECEBIDO PARCIAL"
                    Else
                        TBAbrir!status = "TÍTULO EM ABERTO"
                    End If
                    TBFI.Close
                    TBAbrir!titulodesc = False
                    TBAbrir.Update
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            
            'Fluxo de caixa
            Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IDFluxo
            
            ProcAtualizaLimiteUtil
            '==================================
            Modulo = "Financeiro/Contas à receber/Desconto de duplicata"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Borderô: " & .ListItems(InitFor)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) borderô(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Borderô(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcLimpaValores
    ProcLimpaCamposTotais
    procCarregalistaPrincipal 1
    Lista.ListItems.Clear
    Frame4.Enabled = False
    Frame1.Enabled = False
    cmbBanco.Enabled = False
    txtlocaltroca.Enabled = False
    Novo_Desconto = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirLista()
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
                If USMsgBox("Deseja realmente excluir esta(s) duplicata(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            '==================================
            Modulo = "Financeiro/Contas à receber/Desconto de duplicata"
            Evento = "Excluir duplicata"
            ID_documento = .ListItems(InitFor)
            Documento = "Borderô: " & txtBordero
            Documento1 = "Documento: " & .ListItems(InitFor).ListSubItems(4)
            ProcGravaEvento
            '==================================
            Conexao.Execute "DELETE from troca_titulo_valores where N_conta = " & .ListItems(InitFor)
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber where idintconta = " & .ListItems(InitFor) & " and IDtrocatitulo <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                                
                'Fluxo de Caixa
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBAbrir!IDFluxo), 0, TBAbrir!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    TBFluxo!Operacao = "À Creditar"
                    TBFluxo!Data = TBAbrir!Vencimento
                    TBFluxo!valor = TBAbrir!valor
                    TBFluxo!Instituicao = Null
                    TBFluxo!Hora = Null
                    TBFluxo!status = "N"
                    TBFluxo!Cheque = 0
                    TBFluxo!Bloqueado = False
                    TBFluxo.Update
                End If
                TBFluxo.Close
                TBAbrir!IDtrocatitulo = 0
                
                'Verif. se a conta esta recebida parcial
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from tbl_contas_receber where tituloref = '" & .ListItems(InitFor) & "' and idintconta <> " & .ListItems(InitFor) & " and logsit = 'S'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBAbrir!status = "TÍTULO RECEBIDO PARCIAL"
                Else
                    TBAbrir!status = "TÍTULO EM ABERTO"
                End If
                TBFI.Close
                TBAbrir!titulodesc = False
                TBAbrir.Update
            End If
            TBAbrir.Close
            
            'Atualiza número de duplicatas e prazo médio
            PrazoMedio = 0
            Valorenviado = 0
            Set TBReceber = CreateObject("adodb.recordset")
            TBReceber.Open "Select NDuplicata, Pmedio from troca_titulo where ID = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
            If TBReceber.EOF = False Then
                PrazoMedio = 0
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from troca_titulo_valores where IDduplicata = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    Do While TBContas.EOF = False
                        PrazoMedio = PrazoMedio + (IIf(IsNull(TBContas!valor_enviado), 0, TBContas!valor_enviado) * IIf(IsNull(TBContas!Prazo), 0, TBContas!Prazo))
                        Valorenviado = Valorenviado + IIf(IsNull(TBContas!valor_enviado), 0, TBContas!valor_enviado)
                        TBContas.MoveNext
                    Loop
                End If
                TBContas.Close
                TBReceber!NDuplicata = IIf(IsNull(TBReceber!NDuplicata), 0, TBReceber!NDuplicata) - 1
                If PrazoMedio <> 0 And Valorenviado <> 0 Then PrazoMedio = PrazoMedio / Valorenviado
                TBReceber!Pmedio = PrazoMedio
                TBReceber.Update
            End If
            TBReceber.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) duplicata(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Duplicata(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposTotais
    ProcCarregaLista
    
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from troca_titulo where id = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        
        'Fluxo de caixa
        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
        
        ProcSubtraiVlrResgatadoAnterior
    End If
    TBContas.Close
    
    ProcGravarTotais
    ProcAtualizaLimiteUtil
    ProcLimpaValores
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from troca_titulo where ID = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcCarregaTotais
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaStatusDuplicatas()
On Error GoTo tratar_erro

Set TBTipo = CreateObject("adodb.recordset")
TBTipo.Open "Select * from tbl_contas_receber where IDtrocatitulo = " & IIf(txtBordero = "", 0, txtBordero) & " and Status = 'DUPLICATA DESCONTADA LIQUIDADA'", Conexao, adOpenKeyset, adLockOptimistic
If TBTipo.EOF = False Then
    Permitido = False
    USMsgBox ("Não é permitido descontar esta duplicata, pois existe(m) duplicata(s) líquidada(s)."), vbInformation, "CAPRIND v5.0"
    TBTipo.Close
    Exit Sub
End If
TBTipo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoLista()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Desconto = True Then
    USMsgBox ("Salve o borderô antes de localizar a(s) duplicata(s) para desconto"), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtBordero = "" Then
    NomeCampo = "o borderô"
    Acao = "localizar a(s) duplicata(s) para desconto"
    ProcVerificaAcao
    Exit Sub
End If
If txtBordero <> "" Then
    Permitido = True
    ProcVerificaStatusDuplicatas
    If Permitido = False Then Exit Sub
End If
ProcLimpaValores
frm_trocaduplicata_novo.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frm_Trocaduplicata_relatorios.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Desconto = True Then
    If USMsgBox("O borderô ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Desconto = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Desconto = False
Unload Me

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
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoLista
            Case vbKeyF3: ProcSalvarLista
            Case vbKeyF4: ProcExcluirLista
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: CmdDescDuplicata_Click
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_contas_receber where IDIntconta = " & .ListItems(InitFor) & " and (Status = 'DUPLICATA DESCONTADA LIQUIDADA' or Status = 'DUPLICATA DESCONTADA RECOMPRADA')", Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then GoTo Proximo
                TBContas.Close
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
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_contas_receber where IDIntconta = " & .ListItems(InitFor) & " and (Status = 'DUPLICATA DESCONTADA LIQUIDADA' or Status = 'DUPLICATA DESCONTADA RECOMPRADA')", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                USMsgBox ("Não é permitido excluir esta duplicata pois a mesma já está liquídada ou recomprada."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
            TBContas.Close
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
ProcLimpaValores
txtIDConta.Text = Lista.SelectedItem
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from troca_titulo_valores where N_conta = " & txtIDConta, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtpmedio = IIf(IsNull(TBLISTA!Prazo), "", TBLISTA!Prazo)
    txtvalortitulo = Lista.SelectedItem.SubItems(3)
    txtenviado.Text = IIf(IsNull(TBLISTA!valor_enviado), "", Format(TBLISTA!valor_enviado, "###,##0.00"))
    txtPIS.Text = IIf(IsNull(TBLISTA!valor_pis), "", Format(TBLISTA!valor_pis, "###,##0.00"))
    ProcCalculaPIS
    txtCofins.Text = IIf(IsNull(TBLISTA!valor_cofins), "", Format(TBLISTA!valor_cofins, "###,##0.00"))
    ProcCalculaCofins
    txtretido.Text = IIf(IsNull(TBLISTA!valor_retido), "", Format(TBLISTA!valor_retido, "###,##0.00"))
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaPrincipal_Click()
On Error GoTo tratar_erro

If ListaPrincipal.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from troca_titulo where ID = " & ListaPrincipal.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcCarregaDados
    ProcCarregaLista
    ProcCarregaTotais
    Frame4.Enabled = True
    Frame1.Enabled = True
    cmbBanco.Enabled = True
    txtlocaltroca.Enabled = True
End If
TBAbrir.Close
Novo_Desconto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaPrincipal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaPrincipal
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select IDIntconta from tbl_contas_receber where IDtrocatitulo = " & .ListItems(InitFor) & " and (Status = 'DUPLICATA DESCONTADA LIQUIDADA' or Status = 'DUPLICATA DESCONTADA RECOMPRADA')", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBAbrir.Close
                    GoTo Proximo
                End If
                TBAbrir.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaPrincipal, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaPrincipal_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaPrincipal
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IDIntconta from tbl_contas_receber where IDtrocatitulo = " & .ListItems(InitFor) & " and (Status = 'DUPLICATA DESCONTADA LIQUIDADA' or Status = 'DUPLICATA DESCONTADA RECOMPRADA')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                USMsgBox ("Não é permitido excluir este desconto, pois existe(m) duplicata(s) líquidada(s) ou recomprada(s)."), vbInformation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
            TBAbrir.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtBordero = "0" Or txtBordero = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        Cmb_empresa.Visible = True
        cmbBanco.Visible = True
        txtlocaltroca.Visible = True
        Lista.Visible = False
        ListaPrincipal.Visible = True
        If ListaPrincipal.Visible = True Then ListaPrincipal.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        cmbBanco.Visible = False
        txtlocaltroca.Visible = False
        Lista.Visible = True
        ListaPrincipal.Visible = False
        If Lista.Visible = True Then Lista.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_prazo_Change()
On Error GoTo tratar_erro

If Txt_prazo.Text <> "" Then
    VerifNumero = Txt_prazo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_prazo.Text = ""
        Txt_prazo.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTaxaMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCofins_Change()
On Error GoTo tratar_erro

ProcCalculaCofins
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCofins_LostFocus()
On Error GoTo tratar_erro

If txtCofins.Text <> "" Then
    VerifNumero = txtCofins.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCofins.Text = ""
        txtCofins.SetFocus
        Exit Sub
    End If
    txtCofins.Text = Format(txtCofins.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaCofins()
On Error GoTo tratar_erro

valor = IIf(txtCofins = "", 0, txtCofins)
ValorTotal = IIf(txtvalortitulo = "", 0, txtvalortitulo)
txtVlr_Cofins = Format((ValorTotal * valor) / 100, "###,##0.00")
ProcCalculaValorRetido

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtenviado_LostFocus()
On Error GoTo tratar_erro

txtenviado.Text = Format(txtenviado.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtBordero = ""
txtData.Value = Date
cmbBanco.ListIndex = -1
txtSaldo = ""
txtSaldoAtual = ""
txtlocaltroca = ""
txtObs.Text = ""
txtResponsavel = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

txtvlrtotaltitulo.Text = ""
txtvlrtotalenviado.Text = ""
txtvlrtotalretido.Text = ""
txtvlrtotalresgatado.Text = ""
Txt_data_operacao = "__/__/____"
txttaxatotal.Text = ""
Txt_taxa_mes = ""
Txt_prazo = ""
Valores = 0

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

Private Sub txtpis_Change()
On Error GoTo tratar_erro

ProcCalculaPIS
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpis_LostFocus()
On Error GoTo tratar_erro

If txtPIS.Text <> "" Then
    VerifNumero = txtPIS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPIS.Text = ""
        txtPIS.SetFocus
        Exit Sub
    End If
    txtPIS.Text = Format(txtPIS.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaPIS()
On Error GoTo tratar_erro

valor = IIf(txtPIS = "", 0, txtPIS)
ValorTotal = IIf(txtvalortitulo = "", 0, txtvalortitulo)
txtVlr_PIS = Format((ValorTotal * valor) / 100, "###,##0.00")
ProcCalculaValorRetido

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorRetido()
On Error GoTo tratar_erro

valor = IIf(txtVlr_PIS = "", 0, txtVlr_PIS)
ValorTotal = IIf(txtVlr_Cofins = "", 0, txtVlr_Cofins)
txtretido = Format(valor + ValorTotal, "###,##0.00")
ProcCalculaValorEnviado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorEnviado()
On Error GoTo tratar_erro

valor = IIf(txtvalortitulo = "", 0, txtvalortitulo)
ValorTotal = IIf(txtretido = "", 0, txtretido)
txtenviado = Format(valor - ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procCarregalistaPrincipal(Pagina As Integer)
On Error GoTo tratar_erro

If StrSql_Desconto_Duplicata = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaPrincipal.ListItems.Clear
Set TBLISTA_Desconto_Duplicata = CreateObject("adodb.recordset")
TBLISTA_Desconto_Duplicata.Open StrSql_Desconto_Duplicata, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Desconto_Duplicata.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaPrincipal.ListItems.Clear
TBLISTA_Desconto_Duplicata.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Desconto_Duplicata.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Desconto_Duplicata.PageSize
ContadorReg = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Desconto_Duplicata.RecordCount - IIf(Pagina > 1, (TBLISTA_Desconto_Duplicata.PageSize * (Pagina - 1)), 0), TBLISTA_Desconto_Duplicata.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Desconto_Duplicata.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaPrincipal.ListItems
        With ListaPrincipal.ListItems
            .Add , , TBLISTA_Desconto_Duplicata!ID
            .Item(.Count).SubItems(1) = TBLISTA_Desconto_Duplicata!ID
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Desconto_Duplicata!Data), "", Format(TBLISTA_Desconto_Duplicata!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Desconto_Duplicata!Responsavel), "", TBLISTA_Desconto_Duplicata!Responsavel)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Desconto_Duplicata!banco_recebedor), "", TBLISTA_Desconto_Duplicata!banco_recebedor)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Desconto_Duplicata!local_troca), "", TBLISTA_Desconto_Duplicata!local_troca)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Desconto_Duplicata!Vlrtotalresgatado), "", Format(TBLISTA_Desconto_Duplicata!Vlrtotalresgatado, "###,##0.00"))
        End With
    End With
    TBLISTA_Desconto_Duplicata.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Desconto_Duplicata.RecordCount
If TBLISTA_Desconto_Duplicata.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Desconto_Duplicata.PageCount
ElseIf TBLISTA_Desconto_Duplicata.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Desconto_Duplicata.PageCount & " de: " & TBLISTA_Desconto_Duplicata.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Desconto_Duplicata.AbsolutePage - 1 & " de: " & TBLISTA_Desconto_Duplicata.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

ValorTotal = 0
Lista.ListItems.Clear

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from tbl_Contas_RECEBER where idtrocatitulo = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBProduto.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBProduto.EOF = False
        With Lista.ListItems
            .Add , , TBProduto!IDintconta
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!emissao), "", Format(TBProduto!emissao, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBProduto!Vencimento), "", Format(TBProduto!Vencimento, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBProduto!valor), "", Format(TBProduto!valor, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBProduto!txt_ndocumento), "", TBProduto!txt_ndocumento)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBProduto!NFiscal), "", TBProduto!NFiscal)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBProduto!Parcela), "", TBProduto!Parcela)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProduto!Nome_Razao), "", Trim(TBProduto!Nome_Razao))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProduto!Responsavel), "", Trim(TBProduto!Responsavel))
            ValorTotal = ValorTotal + TBProduto!valor
        End With
        TBProduto.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
txtvlrtotaltitulo.Text = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaValores()
On Error GoTo tratar_erro

txtIDConta.Text = 0
txtpmedio = ""
txtvalortitulo = ""
txtenviado.Text = ""
txtPIS = ""
txtCofins = ""
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Impostos", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If TBFIltro!Duplicata = True Then
        txtPIS = "0"
        txtCofins = "0"
    Else
        txtPIS = TBFIltro!PIS
        txtCofins = TBFIltro!Cofins
    End If
End If
txtVlr_PIS.Text = ""
txtVlr_Cofins.Text = ""
TBFIltro.Close
txtretido.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txttaxatotal_Change()
On Error GoTo tratar_erro

If txttaxatotal.Text <> "" Then
    VerifNumero = txttaxatotal.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txttaxatotal.Text = ""
        txttaxatotal.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTaxaMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTaxaMes()
On Error GoTo tratar_erro

Taxa = IIf(txttaxatotal = "", 0, txttaxatotal)
PrazoMedio = IIf(Txt_prazo = "", 0, Txt_prazo)
Txt_taxa_mes = "0,00"

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select Pmedio from troca_titulo where id = " & IIf(txtBordero = "", 0, txtBordero) & " and Pmedio <> 0", Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    If Taxa <> 0 And PrazoMedio <> 0 Then Txt_taxa_mes = Format((Taxa / PrazoMedio) * 30, "###,##0.00")
End If
TBContas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaNDuplicatasPMedio()
On Error GoTo tratar_erro

PrazoMedio = 0
Valorenviado = 0
Set TBReceber = CreateObject("adodb.recordset")
TBReceber.Open "Select NDuplicata, Pmedio from troca_titulo where ID = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
If TBReceber.EOF = False Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from troca_titulo_valores where IDduplicata = " & txtBordero, Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        Do While TBContas.EOF = False
            PrazoMedio = PrazoMedio + (IIf(IsNull(TBContas!valor_enviado), 0, TBContas!valor_enviado) * IIf(IsNull(TBContas!Prazo), 0, TBContas!Prazo))
            Valorenviado = Valorenviado + IIf(IsNull(TBContas!valor_enviado), 0, TBContas!valor_enviado)
            TBContas.MoveNext
        Loop
    End If
    TBContas.Close
    TBReceber!NDuplicata = IIf(IsNull(TBReceber!NDuplicata), 0, TBReceber!NDuplicata) + 1
    If PrazoMedio <> 0 And Valorenviado <> 0 Then PrazoMedio = PrazoMedio / Valorenviado
    TBReceber!Pmedio = PrazoMedio
    TBReceber.Update
End If
TBReceber.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrtotalenviado_Change()
On Error GoTo tratar_erro

If txtvlrtotalenviado.Text <> "" Then
    VerifNumero = txtvlrtotalenviado.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvlrtotalenviado.Text = ""
        txtvlrtotalenviado.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrtotalenviado_LostFocus()
On Error GoTo tratar_erro

txtvlrtotalenviado.Text = Format(txtvlrtotalenviado.Text, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrtotalresgatado_Change()
On Error GoTo tratar_erro

If txtvlrtotalresgatado.Text <> "" Then
    VerifNumero = txtvlrtotalresgatado.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvlrtotalresgatado.Text = ""
        txtvlrtotalresgatado.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotais IIf(txtvlrtotalenviado = "", 0, txtvlrtotalenviado.Text), IIf(txtvlrtotalretido = "", 0, txtvlrtotalretido), IIf(txtvlrtotalresgatado = "", 0, txtvlrtotalresgatado.Text), TemImposto
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaTotais(VlrEnviado As Double, VlrRetido As Double, VlrResgatado As Double, TemImposto As Boolean)
On Error GoTo tratar_erro

If TemImposto = True Then
    txtvlrtotalresgatado = Format(VlrEnviado - VlrRetido, "###,##0.00")
Else
    txtvlrtotalretido = Format(VlrEnviado - VlrResgatado, "###,##0.00")
    VlrRetido = txtvlrtotalretido
End If
If VlrEnviado <> 0 Then txttaxatotal = Format((VlrRetido / VlrEnviado) * 100, "###,##0.00") Else txttaxatotal = "0,00"
ProcAtualizaSaldo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldo()
On Error GoTo tratar_erro

If cmbBanco <> "" Then
    SaldoAtual = IIf(txtSaldoAtual = "", 0, txtSaldoAtual)
    valor = IIf(txtvlrtotalresgatado = "", 0, txtvlrtotalresgatado)
    txtSaldo = Format(SaldoAtual + valor, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrtotalresgatado_LostFocus()
On Error GoTo tratar_erro

txtvlrtotalresgatado.Text = Format(txtvlrtotalresgatado.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrtotaltitulo_Change()
On Error GoTo tratar_erro

If txtvlrtotaltitulo.Text <> "" Then
    VerifNumero = txtvlrtotaltitulo.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvlrtotaltitulo.Text = ""
        txtvlrtotaltitulo.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvlrtotaltitulo_LostFocus()
On Error GoTo tratar_erro

txtvlrtotaltitulo.Text = Format(txtvlrtotaltitulo.Text, "###,##0.00")

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
    Case 8: ProcAtualizar
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoLista
    Case 2: ProcSalvarLista
    Case 3: ProcExcluirLista
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
