VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoque_Localarmaz 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Local de armazenamento"
   ClientHeight    =   10035
   ClientLeft      =   3615
   ClientTop       =   4590
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEstoque_Localarmaz.frx":0000
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   50
      Top             =   9720
      Width           =   15225
      _ExtentX        =   26855
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
   Begin MSComctlLib.ListView Lista_locarmazenamento 
      Height          =   6345
      Left            =   75
      TabIndex        =   10
      Top             =   2730
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   11192
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
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   8762
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Local de armazenamento"
         Object.Width           =   10261
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Validada"
         Object.Width           =   1499
      EndProperty
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17754
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
      TabCaption(0)   =   "Local de armazenamento"
      TabPicture(0)   =   "frmEstoque_Localarmaz.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtid"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Lista de produtos"
      TabPicture(1)   =   "frmEstoque_Localarmaz.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(1)=   "Lista_loc"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "txtID_item"
      Tab(1).Control(4)=   "txtIdproduto"
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtIdproduto 
         Height          =   285
         Left            =   -66780
         TabIndex        =   53
         Text            =   "0"
         Top             =   1470
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtid 
         Height          =   285
         Left            =   6240
         TabIndex        =   51
         Text            =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   44
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
            TabIndex        =   13
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
            Left            =   2730
            TabIndex        =   11
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
            ItemData        =   "frmEstoque_Localarmaz.frx":0044
            Left            =   6960
            List            =   "frmEstoque_Localarmaz.frx":0051
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   187
            Width           =   1965
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   17
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Localarmaz.frx":0071
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
            TabIndex        =   16
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Localarmaz.frx":3815
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
            TabIndex        =   14
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
            TabIndex        =   15
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Localarmaz.frx":731E
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
            TabIndex        =   18
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmEstoque_Localarmaz.frx":B40D
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
         Begin VB.Label Label15 
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
            Left            =   3360
            TabIndex        =   54
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            Left            =   2040
            TabIndex        =   46
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
            Left            =   5610
            TabIndex        =   45
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.TextBox txtID_item 
         Height          =   285
         Left            =   -73830
         TabIndex        =   31
         Text            =   "0"
         Top             =   5160
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1425
         Left            =   -74925
         TabIndex        =   32
         Top             =   1290
         Width           =   15210
         Begin VB.CheckBox chkPadrao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Padrão"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   14175
            TabIndex        =   26
            Top             =   1015
            Width           =   855
         End
         Begin VB.CommandButton cmdfiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2310
            Picture         =   "frmEstoque_Localarmaz.frx":EC99
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Filtrar por código interno."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox txtUN 
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
            Left            =   13620
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Unidade."
            Top             =   955
            Width           =   390
         End
         Begin VB.CommandButton cmditem 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2640
            Picture         =   "frmEstoque_Localarmaz.frx":F0B4
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar produtos."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox txtFamilia 
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
            Left            =   6090
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   375
            Width           =   8940
         End
         Begin VB.TextBox txtCodInterno 
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
            MaxLength       =   50
            TabIndex        =   19
            ToolTipText     =   "Código interno."
            Top             =   375
            Width           =   2115
         End
         Begin VB.TextBox txtDescricao 
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
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   955
            Width           =   13425
         End
         Begin VB.ComboBox cmbRef 
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
            ItemData        =   "frmEstoque_Localarmaz.frx":F1B6
            Left            =   3060
            List            =   "frmEstoque_Localarmaz.frx":F1B8
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Código de referÊncia."
            Top             =   375
            Width           =   3015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Un."
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
            Left            =   13688
            TabIndex        =   37
            Top             =   765
            Width           =   255
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno"
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
            Left            =   622
            TabIndex        =   36
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Família"
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
            Left            =   10320
            TabIndex        =   35
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
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
            Left            =   6480
            TabIndex        =   34
            Top             =   765
            Width           =   825
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Código de referência"
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
            Left            =   3817
            TabIndex        =   33
            Top             =   180
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1425
         Left            =   75
         TabIndex        =   29
         Top             =   1290
         Width           =   15210
         Begin VB.CheckBox chkPadraoOrdem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Padrão ordem"
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
            Left            =   13020
            TabIndex        =   5
            Top             =   300
            Width           =   1485
         End
         Begin VB.CheckBox chkEstoque 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Desconsiderar estoque"
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
            Left            =   13020
            TabIndex        =   6
            Top             =   540
            Width           =   2145
         End
         Begin VB.CommandButton cmdSetor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   5640
            Picture         =   "frmEstoque_Localarmaz.frx":F1BA
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Localizar setor."
            Top             =   955
            Width           =   315
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
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   375
            Width           =   2025
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
            Left            =   9960
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   375
            Width           =   3015
         End
         Begin VB.TextBox txtdata 
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1185
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
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3135
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   4530
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   375
            Width           =   3375
         End
         Begin VB.TextBox txtSetor 
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
            TabIndex        =   7
            ToolTipText     =   "Setor"
            Top             =   955
            Width           =   5445
         End
         Begin VB.TextBox txtembalagem 
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
            Left            =   6060
            TabIndex        =   9
            ToolTipText     =   "Local de armazenamento."
            Top             =   955
            Width           =   8955
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora da validação"
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
            Left            =   8092
            TabIndex        =   43
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            Left            =   10477
            TabIndex        =   42
            Top             =   180
            Width           =   1980
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
            Index           =   5
            Left            =   600
            TabIndex        =   41
            Top             =   180
            Width           =   345
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
            Index           =   9
            Left            =   2490
            TabIndex        =   40
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            Index           =   17
            Left            =   5985
            TabIndex        =   39
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   0
            Left            =   2707
            TabIndex        =   38
            Top             =   765
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local de armazenamento"
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
            Index           =   0
            Left            =   9645
            TabIndex        =   30
            Top             =   765
            Width           =   1785
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   52
         Top             =   330
         Width           =   15210
         _ExtentX        =   26829
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
         ButtonCaption8  =   "Status"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Status (F7)"
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
         ButtonWidth8    =   39
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Validação"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Validar/Cancelar validação (F10)"
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
         ButtonLeft9     =   388
         ButtonTop9      =   2
         ButtonWidth9    =   53
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Atualizar"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft10    =   443
         ButtonTop10     =   2
         ButtonWidth10   =   50
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonAlignment11=   2
         ButtonType11    =   1
         ButtonStyle11   =   -1
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   -1
         ButtonLeft11    =   495
         ButtonTop11     =   4
         ButtonWidth11   =   2
         ButtonHeight11  =   54
         ButtonCaption12 =   "Ajuda"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Ajuda (F1)"
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
         ButtonLeft12    =   499
         ButtonTop12     =   2
         ButtonWidth12   =   41
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonCaption13 =   "Sair"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Sair (Esc)"
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
         ButtonLeft13    =   542
         ButtonTop13     =   2
         ButtonWidth13   =   30
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonKey14     =   "14"
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   574
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   10710
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmEstoque_Localarmaz.frx":F2BC
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_loc 
         Height          =   6975
         Left            =   -74925
         TabIndex        =   27
         Top             =   2730
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12303
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
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "IDproduto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   17824
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   5292
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   945
         Left            =   -74925
         TabIndex        =   49
         Top             =   330
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1667
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
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior"
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
         ButtonLeft5     =   190
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
         ButtonLeft6     =   247
         ButtonTop6      =   4
         ButtonWidth6    =   2
         ButtonHeight6   =   52
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   251
         ButtonTop7      =   2
         ButtonWidth7    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   289
         ButtonTop8      =   2
         ButtonWidth8    =   26
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonEnabled9  =   0   'False
         ButtonKey9      =   "9"
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
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   10710
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmEstoque_Localarmaz.frx":173EE
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmEstoque_Localarmaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_LocalArmaz             As Boolean 'OK
Public Novo_LocalArmaz2            As Boolean 'OK
Public Sql_localarmaz_Localizar As String 'OK
Dim TBLISTA_LocalArmaz As ADODB.Recordset 'OK
Public FormulaRel_Local As String

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista_locarmazenamento.ListItems.Count = 0 Then Exit Sub
NomeRel = "Estoque_LocalArmazenamento.rpt"

ProcImprimirRel "", ""

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
With Lista_locarmazenamento
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(ais) de armazenamento antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmEstoque_Localarmaz_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Item()
On Error GoTo tratar_erro

Lista_loc.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select EL.ID, EL.idproduto, EL.Codinterno, P.Descricao, P.classe from Estoque_Localarmazenamento EL INNER JOIN projproduto P ON P.codproduto = EL.idproduto where EL.idemb_locarm = " & txtId & " order by EL.Codinterno", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBAbrir.MoveFirst
    Do While TBAbrir.EOF = False
        With Lista_loc.ListItems
            .Add , , TBAbrir!ID
            .Item(.Count).SubItems(1) = TBAbrir!IDProduto
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Classe), "", TBAbrir!Classe)
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

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista_locarmazenamento.ListItems.Clear
If Sql_localarmaz_Localizar = "" Then Exit Sub
Set TBLISTA_LocalArmaz = CreateObject("adodb.recordset")
TBLISTA_LocalArmaz.Open Sql_localarmaz_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_LocalArmaz.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista_locarmazenamento.ListItems.Clear
TBLISTA_LocalArmaz.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_LocalArmaz.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_LocalArmaz.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_LocalArmaz.RecordCount - IIf(Pagina > 1, (TBLISTA_LocalArmaz.PageSize * (Pagina - 1)), 0), TBLISTA_LocalArmaz.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_LocalArmaz.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_locarmazenamento.ListItems.Add(, , TBLISTA_LocalArmaz!ID)
        .SubItems(1) = IIf(IsNull(TBLISTA_LocalArmaz!Data), "", Format(TBLISTA_LocalArmaz!Data, "dd/mm/yy"))
        .SubItems(2) = IIf(IsNull(TBLISTA_LocalArmaz!Responsavel), "", TBLISTA_LocalArmaz!Responsavel)
        .SubItems(3) = IIf(IsNull(TBLISTA_LocalArmaz!Setor), "", TBLISTA_LocalArmaz!Setor)
        .SubItems(4) = IIf(IsNull(TBLISTA_LocalArmaz!Descricao), "", TBLISTA_LocalArmaz!Descricao)
        If IsNull(TBLISTA_LocalArmaz!DtValidacao) = False And TBLISTA_LocalArmaz!DtValidacao <> "" Then .SubItems(5) = "SIM" Else .SubItems(5) = "NÃO"
        
        If TBLISTA_LocalArmaz!Descricao = "SERVIÇOS" Or TBLISTA_LocalArmaz!Descricao = "RETORNO DE MERCADORIA" Or TBLISTA_LocalArmaz!Descricao = "INDUSTRIALIZAÇÃO" Or TBLISTA_LocalArmaz!Descricao = "ESTOQUE PADRÃO" Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbBlue
            .ListSubItems(2).ForeColor = vbBlue
            .ListSubItems(3).ForeColor = vbBlue
            .ListSubItems(4).ForeColor = vbBlue
            .ListSubItems(5).ForeColor = vbBlue
        End If
    End With
    TBLISTA_LocalArmaz.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_LocalArmaz.RecordCount
If TBLISTA_LocalArmaz.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_LocalArmaz.PageCount
ElseIf TBLISTA_LocalArmaz.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_LocalArmaz.PageCount & " de: " & TBLISTA_LocalArmaz.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_LocalArmaz.AbsolutePage - 1 & " de: " & TBLISTA_LocalArmaz.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

frmEstoque_Localarmaz_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Estoque_Localarmazenamento_criar order by descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("ID = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId = TBLISTA!ID
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_Localarmazenamento_criar where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
        ProcCarregaLista_Item
    Else
        USMsgBox ("Fim dos cadastros de locais de armazenamento."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_LocalArmaz = False
Novo_LocalArmaz2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362L" Then
    If USMsgBox("Deseja realmente atualizar os dados dos locais de armazenamento?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "select * from Estoque_Localarmazenamento", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            PBLista.Max = TBAbrir.RecordCount
            PBLista.Min = 0
            PBLista.Value = 0
            Contador = 0
            TBAbrir.MoveFirst
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!Codinterno) = False Then
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "select * from projproduto where desenho = '" & TBAbrir!Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        TBAbrir!IDProduto = IIf(IsNull(TBItem!Codproduto), "0", TBItem!Codproduto)
                        TBAbrir.Update
                    End If
                    TBItem.Close
                End If
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
    
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Local de armazenamento"
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

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_locarmazenamento
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) local(ais) de armazenamento?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Estoque_Localarmazenamento where idemb_locarm = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Estoque_Localarmazenamento_criar where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Estoque/Local de armazenamento"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Local de armazenamento: " & .ListItems(InitFor).ListSubItems(4)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(ais) de armazenamento antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Local(ais) de armazenamento excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Novo_LocalArmaz = False
    Frame2.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_item()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_loc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Estoque_Localarmazenamento where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Estoque/Local de armazenamento"
            Evento = "Excluir produto"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & .ListItems(InitFor).ListSubItems(2)
            Documento1 = "Local de armazenamento: " & Lista_locarmazenamento.SelectedItem.ListSubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposItem
    ProcCarregaLista_Item
    Novo_LocalArmaz2 = False
    Frame4.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lista_locarmazenamento
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(8) = 5
        .ButtonState(9) = 5
    ElseIf Cmb_opcao_lista = "Validação" Then
        .ButtonState(4) = 5
        .ButtonState(8) = 5
        .ButtonState(9) = 0
        Else
            .ButtonState(4) = 5
            .ButtonState(8) = 0
            .ButtonState(9) = 5
        End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

If txtCodinterno = "" Then Exit Sub
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "select * from projproduto where Desenho = '" & txtCodinterno & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtCodinterno = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
    ProcCarregaComboCodRef cmbRef, "P.codproduto = " & TBItem!Codproduto, 0, "", False, True
    txtfamilia = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
    txtdescricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
    txtUN = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
    txtidproduto = TBItem!Codproduto
End If
TBItem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmditem_Click()
On Error GoTo tratar_erro

frmEstoque_Localarmaz_item.Show 1

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
cmdSetor_Click
Novo_LocalArmaz = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_item()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "local de armazenamento", "produto", True) = False Then Exit Sub
ProcLimpaCamposItem
Frame4.Enabled = True
txtCodinterno.SetFocus
Novo_LocalArmaz2 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Estoque_Localarmazenamento_criar order by descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.Find ("ID = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId = TBLISTA!ID
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_Localarmazenamento_criar where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
        ProcCarregaLista_Item
    Else
        USMsgBox ("Fim dos cadastros de locais de armazenamento."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_LocalArmaz = False
Novo_LocalArmaz2 = False

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
If txtembalagem = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    txtembalagem.SetFocus
    Exit Sub
End If
If Novo_LocalArmaz = False And (txtembalagem = "SERVIÇOS" Or txtembalagem = "RETORNO DE MERCADORIA" Or txtembalagem = "INDUSTRIALIZAÇÃO" Or txtembalagem = "ESTOQUE PADRÃO") Then
    USMsgBox ("Não é permitido alterar, pois o mesmo é um local de armazenamento padrão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Estoque_Localarmazenamento_criar where Descricao = '" & txtembalagem & "' and id <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Já existe um local de armazenamento com este nome, favor alterar."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If

Permitido = True
If chkPadraoOrdem.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select Descricao, PadraoOrdem from Estoque_Localarmazenamento_Criar where id <> " & txtId & " and PadraoOrdem = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If USMsgBox("O local de armazenamento " & TBAbrir!Descricao & " esta como padrão da ordem, deseja alterar?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            Permitido = False
            chkPadraoOrdem.Value = 0
        Else
            TBAbrir!PadraoOrdem = False
            TBAbrir.Update
        End If
    End If
    TBAbrir.Close
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Estoque_Localarmazenamento_criar where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    TBAbrir.AddNew
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "este local de armazenamento", True) = False Then Exit Sub
    If TBAbrir!Descricao <> txtembalagem Then
        Conexao.Execute "Update estoque_controle Set Local_armaz = '" & txtembalagem & "' where Local_armaz = '" & TBAbrir!Descricao & "'"
        Conexao.Execute "Update Estoque_controle_recebimento Set Local_armaz = '" & txtembalagem & "' where Local_armaz = '" & TBAbrir!Descricao & "'"
        Conexao.Execute "Update Estoque_fisico Set Local_armaz = '" & txtembalagem & "' where Local_armaz = '" & TBAbrir!Descricao & "'"
    End If
End If
TBAbrir!Data = IIf(txtData = "", Date, txtData)
TBAbrir!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
TBAbrir!status = IIf(txtStatus = "", "Liberado", txtStatus)
TBAbrir!Setor = txtSetor
TBAbrir!Descricao = txtembalagem
If chkEstoque.Value = 1 Then TBAbrir!Estoque = True Else TBAbrir!Estoque = False
If chkPadraoOrdem.Value = 1 And Permitido = True Then TBAbrir!PadraoOrdem = True Else TBAbrir!PadraoOrdem = False
TBAbrir.Update
txtId = TBAbrir!ID
TBAbrir.Close
If Novo_LocalArmaz = True Then
    USMsgBox ("Novo local de armazenamento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_localarmaz_Localizar = "Select * from Estoque_Localarmazenamento_criar where ID = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista_locarmazenamento.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista_locarmazenamento.SelectedItem = Lista_locarmazenamento.ListItems(CodigoLista)
        Lista_locarmazenamento.SetFocus
    End If
End If
'==================================
Modulo = "Estoque/Local de armazenamento"
ID_documento = txtId
Documento = "Local de armazenamento: " & txtembalagem
Documento1 = ""
ProcGravaEvento
'=================================
Novo_LocalArmaz = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_item()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "local de armazenamento", "o produto", True) = False Then Exit Sub
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtidproduto = 0 Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    frmEstoque_Localarmaz_item.Show 1
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select codinterno from Estoque_Localarmazenamento where codinterno = '" & txtCodinterno & "' and idemb_locarm = " & txtId & " and id <> " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("O produto " & txtCodinterno & " já foi cadastrado para este local de armazenamento."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close
Permitido = True
If chkPadrao.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select ELC.Descricao, EL.Padrao from Estoque_Localarmazenamento EL INNER JOIN Estoque_Localarmazenamento_Criar ELC on ELC.ID = EL.idemb_locarm where EL.codinterno = '" & txtCodinterno & "' and EL.idemb_locarm <> " & txtId & " and EL.Padrao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If USMsgBox("O produto já possui o local de armazenamento " & TBAbrir!Descricao & " como padrão, deseja alterar?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            Permitido = False
            chkPadrao.Value = 0
        Else
            TBAbrir!Padrao = False
            TBAbrir.Update
        End If
    End If
    TBAbrir.Close
End If

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from Estoque_Localarmazenamento where id = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = True Then
    TBItem.AddNew
    TBItem!idemb_locarm = txtId
End If
TBItem!IDProduto = txtidproduto
TBItem!Codinterno = txtCodinterno
If chkPadrao.Value = 1 And Permitido = True Then TBItem!Padrao = True Else TBItem!Padrao = False
TBItem.Update
txtID_item = TBItem!ID
TBItem.Close
ProcCarregaLista_Item
If Novo_LocalArmaz2 = True Then
    USMsgBox ("Novo produto cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto"
    If CodigoLista1 <> 0 And Lista_loc.ListItems.Count <> 0 Then
        Lista_loc.SelectedItem = Lista_loc.ListItems(CodigoLista1)
        Lista_loc.SetFocus
    End If
End If
'==================================
Modulo = "Estoque/Local de armazenamento"
ID_documento = txtID_item
Documento = "Local de armazenamento: " & txtembalagem
Documento1 = "Código interno: " & txtCodinterno
ProcGravaEvento
'=================================
Novo_LocalArmaz2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_LocalArmaz.AbsolutePage <> 2 Then
    If TBLISTA_LocalArmaz.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_LocalArmaz.PageCount - 1)
    Else
        TBLISTA_LocalArmaz.AbsolutePage = TBLISTA_LocalArmaz.AbsolutePage - 2
        ProcExibePagina (TBLISTA_LocalArmaz.AbsolutePage)
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
    TBLISTA_LocalArmaz.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_LocalArmaz.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_LocalArmaz.AbsolutePage = 1
ProcExibePagina (TBLISTA_LocalArmaz.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_LocalArmaz.AbsolutePage <> -3 Then
    If TBLISTA_LocalArmaz.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_LocalArmaz.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_LocalArmaz.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_LocalArmaz.AbsolutePage = TBLISTA_LocalArmaz.PageCount
ProcExibePagina (TBLISTA_LocalArmaz.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSetor_Click()
On Error GoTo tratar_erro

CadMaquinas = False
Funcionario = False
Usuarios = False
Estoque_Local_Armazenamento = True
frmUsuarios_Setor.Show 1

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
            Case vbKeyF2: ProcAbrir
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF7: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF10: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista_locarmazenamento, "Estoque/Local de armazenamento"
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_item
            Case vbKeyF3: procSalvar_item
            Case vbKeyF4: procExcluir_item
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 14, True
ProcCarregaToolBar2 Me, 15195, 8, True
Formulario = "Estoque/Local de armazenamento"
Direitos
SSTab1.Tab = 0
Cmb_opcao_lista.Text = "Validação"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_LocalArmaz = True Then
    If USMsgBox("O local de armazenamento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_LocalArmaz = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_LocalArmaz2 = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_item
        If Novo_LocalArmaz2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_LocalArmaz = False
Novo_LocalArmaz2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_loc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_loc
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Estoque_Localarmazenamento_criar", "ID = " & txtId, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_loc, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_loc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_loc
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Estoque_Localarmazenamento_criar", "ID = " & txtId, "o local de armazenamento", "produto(s)", "excluir este(s)", True, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_loc_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_loc.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from Estoque_Localarmazenamento where id = " & Lista_loc.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposItem
    procPuxadadosItem
    CodigoLista1 = Lista_loc.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId.Text = 0
txtData.Text = Format(Date, "dd/mm/yy")
txtResponsavel.Text = pubUsuario
txtStatus.Text = "Liberado"
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtSetor.Text = ""
txtembalagem = ""
chkEstoque.Value = 0
chkPadraoOrdem.Value = 0
ProcMostrarEsconderTab txtembalagem
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_locarmazenamento_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_locarmazenamento
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems(InitFor).ListSubItems(4) = "SERVIÇOS" Or .ListItems(InitFor).ListSubItems(4) = "RETORNO DE MERCADORIA" Or .ListItems(InitFor).ListSubItems(4) = "INDUSTRIALIZAÇÃO" Or .ListItems(InitFor).ListSubItems(4) = "ESTOQUE PADRÃO" Then GoTo Proximo
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("Estoque_Localarmazenamento_criar", "id = " & .ListItems.Item(InitFor), True) = False Then GoTo Proximo
                    ProcVerificaRegistroUtilizadoSemMsg "Estoque_Controle", "local_armaz = '" & .ListItems(InitFor).ListSubItems(4) & "'"
                    If Permitido = False Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_locarmazenamento, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_locarmazenamento_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_locarmazenamento
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If .ListItems(InitFor).ListSubItems(4) = "SERVIÇOS" Or .ListItems(InitFor).ListSubItems(4) = "RETORNO DE MERCADORIA" Or .ListItems(InitFor).ListSubItems(4) = "INDUSTRIALIZAÇÃO" Or .ListItems(InitFor).ListSubItems(4) = "ESTOQUE PADRÃO" Then
                If Cmb_opcao_lista = "Status" Then
                    MsgTexto = "alterar status"
                ElseIf Cmb_opcao_lista = "Validação" Then
                        MsgTexto = "cancelar validação"
                    Else
                        MsgTexto = "excluir"
                End If
                USMsgBox ("Não é permitido " & MsgTexto & ", pois o mesmo é um local de armazenamento padrão."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("Estoque_Localarmazenamento_criar", "id = " & .ListItems.Item(InitFor), "o mesmo", "local de armazenamento", "excluir este", True, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                Mensagem = "Não é permitido excluir este local de armazenamento, pois o mesmo está sendo utilizado no módulo"
                ProcVerificaRegistroUtilizado "Estoque_Controle", "local_armaz = '" & .ListItems(InitFor).ListSubItems(4) & "'", "Estoque/Movimentacao"
                If Permitido = False Then
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

Private Sub Lista_locarmazenamento_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_locarmazenamento.ListItems.Count = 0 Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Estoque_Localarmazenamento_criar where ID = " & Lista_locarmazenamento.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista_locarmazenamento.SelectedItem.index
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId = TBProduto!ID
txtembalagem = TBProduto!Descricao
txtData = IIf(IsNull(TBProduto!Data), "", TBProduto!Data)
txtResponsavel = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)
txtStatus = IIf(IsNull(TBProduto!status), "", TBProduto!status)
txtSetor = IIf(IsNull(TBProduto!Setor), "", TBProduto!Setor)
txtDtValidacao = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao)
txtRespValidacao = IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao)
If TBProduto!Estoque = True Then chkEstoque.Value = 1 Else chkEstoque.Value = 0
If TBProduto!PadraoOrdem = True Then chkPadraoOrdem.Value = 1 Else chkPadraoOrdem.Value = 0
Caption = "Estoque - Local de armazenamento (Descrição : " & TBProduto!Descricao & ")"
Frame2.Enabled = True
Novo_LocalArmaz = False
ProcLimparTudo

ProcMostrarEsconderTab TBProduto!Descricao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarEsconderTab(Descricao As String)
On Error GoTo tratar_erro

With SSTab1
    If Descricao = "SERVIÇOS" Or Descricao = "RETORNO DE MERCADORIA" Or Descricao = "INDUSTRIALIZAÇÃO" Or Descricao = "ESTOQUE PADRÃO" Then
        .TabVisible(1) = False
        .TabsPerRow = 1
    Else
        .TabVisible(1) = True
        .TabsPerRow = 2
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame4.Enabled = False
ProcLimpaCamposItem
Novo_LocalArmaz2 = False

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
Select Case SSTab1.Tab
    Case 0:
        Lista_loc.Visible = False
        With Lista_locarmazenamento
            .Visible = True
            .SetFocus
        End With
    Case 1:
        Lista_locarmazenamento.Visible = False
        With Lista_loc
            .Visible = True
            .SetFocus
        End With
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcLimpaCamposItem
        ProcCarregaLista_Item
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_LocalArmaz = True Then
    USMsgBox ("Salve o local de armazenamento antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposItem()
On Error GoTo tratar_erro

txtID_item = 0
txtCodinterno = ""
cmbRef.Clear
txtfamilia = ""
txtdescricao = ""
txtUN = ""
chkPadrao.Value = 0
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxadadosItem()
On Error GoTo tratar_erro

txtID_item = TBAbrir!ID
txtCodinterno = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
txtidproduto = IIf(IsNull(TBAbrir!IDProduto), "0", TBAbrir!IDProduto)
If TBAbrir!Padrao = True Then chkPadrao.Value = 1 Else chkPadrao.Value = 0
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "select * from projproduto where codproduto = " & txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtCodinterno = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
    ProcCarregaComboCodRef cmbRef, "P.codproduto = " & txtidproduto, 0, "", False, True
    txtfamilia = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
    txtdescricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
    txtUN = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
End If
Novo_LocalArmaz2 = False
Frame4.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodInterno_Change()
On Error GoTo tratar_erro

txtidproduto = 0
cmbRef.Clear
txtfamilia = ""
txtdescricao = ""
txtUN = ""
    
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
    Case 2: ProcAbrir
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcStatus
    Case 9: ProcValidarRegistros Lista_locarmazenamento, "Estoque/Local de armazenamento"
    Case 10: procAtualiza
    'Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovo_item
    Case 2: procSalvar_item
    Case 3: procExcluir_item
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
