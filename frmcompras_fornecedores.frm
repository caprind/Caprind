VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_fornecedores 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Fornecedores"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
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
   Icon            =   "frmcompras_fornecedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   975
      Left            =   30
      TabIndex        =   276
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
      ButtonLeft8     =   313
      ButtonTop8      =   2
      ButtonWidth8    =   41
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Sair"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Sair (ESC)"
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
      ButtonLeft9     =   356
      ButtonTop9      =   2
      ButtonWidth9    =   30
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonKey10     =   "10"
      ButtonAlignment10=   2
      BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11460
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   13230
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmcompras_fornecedores.frx":014A
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4890
      Left            =   75
      TabIndex        =   39
      Top             =   4500
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8625
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
         Object.Width           =   512
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
         Text            =   "Razão social"
         Object.Width           =   19015
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Validado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   278
      Top             =   9360
      Width           =   15195
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
         TabIndex        =   40
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
         TabIndex        =   42
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   46
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcompras_fornecedores.frx":552E
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
         TabIndex        =   45
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcompras_fornecedores.frx":8CD2
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
         TabIndex        =   43
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
         TabIndex        =   44
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcompras_fornecedores.frx":C7DB
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
         TabIndex        =   47
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmcompras_fornecedores.frx":108CA
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
         ItemData        =   "frmcompras_fornecedores.frx":14156
         Left            =   6960
         List            =   "frmcompras_fornecedores.frx":14163
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Index           =   48
         Left            =   3360
         TabIndex        =   307
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
         Index           =   2
         Left            =   5610
         TabIndex        =   288
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label25 
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
         TabIndex        =   282
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
         TabIndex        =   280
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
         TabIndex        =   279
         Top             =   240
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   134
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
      TabCaption(0)   =   "Fornecedor"
      TabPicture(0)   =   "frmcompras_fornecedores.frx":14183
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtIDCliente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USToolBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contatos"
      TabPicture(1)   =   "frmcompras_fornecedores.frx":1419F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIDContato"
      Tab(1).Control(1)=   "Lista_contato"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Entrega"
      TabPicture(2)   =   "frmcompras_fornecedores.frx":141BB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListaEntrega"
      Tab(2).Control(1)=   "txtid_entrega"
      Tab(2).Control(2)=   "Frame11"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Cobrança"
      TabPicture(3)   =   "frmcompras_fornecedores.frx":141D7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtid_cobranca"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "listacobranca"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame16"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Famílias"
      TabPicture(4)   =   "frmcompras_fornecedores.frx":141F3
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(1)=   "txtid_familia"
      Tab(4).Control(2)=   "lista_familia"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Dados bancário"
      TabPicture(5)   =   "frmcompras_fornecedores.frx":1420F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame10"
      Tab(5).Control(1)=   "txtid_banco"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lista_banco"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Segmentos"
      TabPicture(6)   =   "frmcompras_fornecedores.frx":1422B
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtID_segmento"
      Tab(6).Control(1)=   "Frame5"
      Tab(6).Control(2)=   "Lista_Segmento"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Outros"
      TabPicture(7)   =   "frmcompras_fornecedores.frx":14247
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame18"
      Tab(7).Control(1)=   "USToolBar3"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Comercial"
      TabPicture(8)   =   "frmcompras_fornecedores.frx":14263
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame4"
      Tab(8).ControlCount=   1
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   45
         TabIndex        =   275
         Top             =   330
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   1720
         ButtonCount     =   15
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
         ButtonToolTipText8=   "Carregar/Limpar lista de fornecedores cadastrados (F7)"
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
         ButtonCaption9  =   "Status"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Status (F8)"
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
         ButtonLeft9     =   432
         ButtonTop9      =   2
         ButtonWidth9    =   45
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Validação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Validar/Cancelar validação (F10)"
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
         ButtonWidth10   =   62
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Atualizar"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Usado pelo administrador do sistema"
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
         ButtonLeft11    =   543
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
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState12   =   -1
         ButtonLeft12    =   604
         ButtonTop12     =   4
         ButtonWidth12   =   2
         ButtonHeight12  =   54
         ButtonCaption13 =   "Ajuda"
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonToolTipText13=   "Ajuda"
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
         ButtonLeft13    =   608
         ButtonTop13     =   2
         ButtonWidth13   =   41
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonCaption14 =   "Sair"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Sair (ESC)"
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
         ButtonLeft14    =   651
         ButtonTop14     =   2
         ButtonWidth14   =   30
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
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
         ButtonState15   =   5
         ButtonLeft15    =   683
         ButtonTop15     =   2
         ButtonWidth15   =   24
         ButtonHeight15  =   24
         ButtonUseMaskColor15=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   13050
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmcompras_fornecedores.frx":1427F
            Count           =   1
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7815
         Left            =   -74925
         TabIndex        =   310
         Top             =   1285
         Width           =   15255
         Begin VB.TextBox txtID_cfop 
            Alignment       =   2  'Center
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
            Left            =   7770
            MaxLength       =   100
            TabIndex        =   327
            ToolTipText     =   "Data da revisão."
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox txtCFOP 
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
            Left            =   7185
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   328
            TabStop         =   0   'False
            ToolTipText     =   "Natureza da operação."
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtoperacao 
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
            Left            =   8265
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   326
            TabStop         =   0   'False
            ToolTipText     =   "Descrição da natureza da operação."
            Top             =   240
            Width           =   6105
         End
         Begin VB.CommandButton cmdcfop 
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
            Left            =   14370
            Picture         =   "frmcompras_fornecedores.frx":1CB4A
            Style           =   1  'Graphical
            TabIndex        =   325
            ToolTipText     =   "Localizar CFOP."
            Top             =   240
            Width           =   315
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
            ItemData        =   "frmcompras_fornecedores.frx":1CC4C
            Left            =   2235
            List            =   "frmcompras_fornecedores.frx":1CC4E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   324
            ToolTipText     =   "Empresa."
            Top             =   240
            Width           =   4245
         End
         Begin VB.CommandButton Cmd_limpar_CFOP 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmcompras_fornecedores.frx":1CC50
            Style           =   1  'Graphical
            TabIndex        =   323
            ToolTipText     =   "Limpar grupo."
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtimpostos 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   319
            ToolTipText     =   "Impostos."
            Top             =   2880
            Width           =   12465
         End
         Begin VB.TextBox txtValidade 
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
            Height          =   1130
            Left            =   2235
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   318
            ToolTipText     =   "Prazo de validade da proposta."
            Top             =   4020
            Width           =   12465
         End
         Begin VB.TextBox txttransporte 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   317
            ToolTipText     =   "Transporte."
            Top             =   1740
            Width           =   12465
         End
         Begin VB.CommandButton cmdTransporte_padrao 
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
            Height          =   1130
            Left            =   14700
            Picture         =   "frmcompras_fornecedores.frx":1CD8E
            Style           =   1  'Graphical
            TabIndex        =   316
            ToolTipText     =   "Localizar transporte."
            Top             =   1740
            Width           =   315
         End
         Begin VB.CommandButton cmdImpostos_padrao 
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
            Height          =   1130
            Left            =   14700
            Picture         =   "frmcompras_fornecedores.frx":1CE90
            Style           =   1  'Graphical
            TabIndex        =   315
            ToolTipText     =   "Localizar impostos."
            Top             =   2880
            Width           =   315
         End
         Begin VB.CommandButton cmdValidade_Padrao 
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
            Height          =   1130
            Left            =   14700
            Picture         =   "frmcompras_fornecedores.frx":1CF92
            Style           =   1  'Graphical
            TabIndex        =   314
            ToolTipText     =   "Localizar validade."
            Top             =   4020
            Width           =   315
         End
         Begin VB.TextBox txtCondicoes 
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
            Height          =   1130
            Left            =   2235
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   312
            ToolTipText     =   "Condições de pagamento."
            Top             =   600
            Width           =   12465
         End
         Begin VB.CommandButton cmdCond_pag_padrao 
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
            Height          =   1130
            Left            =   14700
            Picture         =   "frmcompras_fornecedores.frx":1D094
            Style           =   1  'Graphical
            TabIndex        =   311
            ToolTipText     =   "Localizar cond. de pagamento."
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP :"
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
            Height          =   165
            Left            =   6600
            TabIndex        =   330
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa:"
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
            Left            =   1380
            TabIndex        =   329
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Impostos :"
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
            Left            =   1395
            TabIndex        =   322
            Top             =   2880
            Width           =   765
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Validade :"
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
            Left            =   1455
            TabIndex        =   321
            Top             =   4020
            Width           =   705
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte :"
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
            Left            =   1260
            TabIndex        =   320
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Condições de pagamento :"
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
            Left            =   240
            TabIndex        =   313
            Top             =   600
            Width           =   1920
         End
      End
      Begin VB.TextBox txtID_segmento 
         Height          =   315
         Left            =   -71490
         TabIndex        =   304
         Text            =   "0"
         Top             =   3720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   300
         Top             =   1285
         Width           =   15200
         Begin VB.CommandButton cmdSegmento 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmcompras_fornecedores.frx":1D196
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Localizar segmento."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtSegmento 
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
            Left            =   4410
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   113
            TabStop         =   0   'False
            ToolTipText     =   "Segmento."
            Top             =   390
            Width           =   10245
         End
         Begin VB.TextBox txtResp_segmento 
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
            TabIndex        =   112
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3015
         End
         Begin VB.TextBox txtData_segmento 
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
            TabIndex        =   111
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label2 
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
            Index           =   21
            Left            =   2430
            TabIndex        =   303
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   19
            Left            =   600
            TabIndex        =   302
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Segmento"
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
            Index           =   43
            Left            =   9172
            TabIndex        =   301
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   -74925
         TabIndex        =   295
         Top             =   1290
         Width           =   15195
         Begin VB.TextBox Txt_ICMS_ind 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   308
            ToolTipText     =   "Alíquota de ICMS para industrialização."
            Top             =   1695
            Width           =   1185
         End
         Begin VB.CheckBox chkICMSST 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Utiliza cálculo simplificado de ICMS ST"
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
            Left            =   5070
            TabIndex        =   128
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CheckBox chkDesignado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Designado/homologado/aprovado por cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   195
            TabIndex        =   126
            Top             =   1200
            Width           =   3525
         End
         Begin VB.CheckBox chkAprovada_Cliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fonte aprovada pelo cliente"
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
            TabIndex        =   118
            Top             =   720
            Width           =   2325
         End
         Begin VB.CheckBox Chk_certificado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Certificado"
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
            TabIndex        =   116
            Top             =   240
            Width           =   1095
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo/caminho do arquivo"
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
            Height          =   735
            Left            =   4260
            TabIndex        =   306
            Top             =   180
            Width           =   10725
            Begin VB.CommandButton cmdImportar 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   9570
               Picture         =   "frmcompras_fornecedores.frx":1D298
               Style           =   1  'Graphical
               TabIndex        =   123
               ToolTipText     =   "Localizar arquivo."
               Top             =   270
               Width           =   315
            End
            Begin VB.TextBox Txt_caminho_certificado 
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
               Left            =   2790
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   122
               TabStop         =   0   'False
               ToolTipText     =   "Caminho do arquivo."
               Top             =   270
               Width           =   6765
            End
            Begin VB.CommandButton Cmd_limpar_caminho 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   9900
               Picture         =   "frmcompras_fornecedores.frx":1D39A
               Style           =   1  'Graphical
               TabIndex        =   124
               ToolTipText     =   "Limpar caminho."
               Top             =   270
               Width           =   315
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   10230
               Picture         =   "frmcompras_fornecedores.frx":1D4D8
               Style           =   1  'Graphical
               TabIndex        =   125
               ToolTipText     =   "Visualizar arquivo."
               Top             =   270
               Width           =   315
            End
            Begin VB.CommandButton cmdCertificado 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2370
               Picture         =   "frmcompras_fornecedores.frx":1DA9A
               Style           =   1  'Graphical
               TabIndex        =   121
               ToolTipText     =   "Localizar tipo do arquivo."
               Top             =   270
               Width           =   315
            End
            Begin VB.TextBox txtCertificado 
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
               TabIndex        =   120
               ToolTipText     =   "Tipo do arquivo."
               Top             =   270
               Width           =   2175
            End
         End
         Begin VB.CommandButton Cmd_localizar_tipo_dcto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1230
            Picture         =   "frmcompras_fornecedores.frx":1DB9C
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Localizar tipo do documento."
            Top             =   1695
            Width           =   315
         End
         Begin VB.ComboBox cmbTipo_doc 
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
            ItemData        =   "frmcompras_fornecedores.frx":1DC9E
            Left            =   180
            List            =   "frmcompras_fornecedores.frx":1DCA0
            Style           =   2  'Dropdown List
            TabIndex        =   129
            ToolTipText     =   "Tipo do documento previsto para pagamento."
            Top             =   1695
            Width           =   1035
         End
         Begin VB.CheckBox chkSedex 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sedex"
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
            Height          =   225
            Left            =   3982
            TabIndex        =   127
            Top             =   1200
            Width           =   825
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
            Height          =   885
            Left            =   8220
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   132
            ToolTipText     =   "Observações."
            Top             =   1140
            Width           =   6765
         End
         Begin VB.CheckBox Chk_avaliado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Avaliado"
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
            TabIndex        =   117
            Top             =   480
            Width           =   915
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
            ItemData        =   "frmcompras_fornecedores.frx":1DCA2
            Left            =   1680
            List            =   "frmcompras_fornecedores.frx":1DCA4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   131
            ToolTipText     =   "Instituição bancária prevista para recebimento."
            Top             =   1695
            Width           =   5265
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dt. vencimento"
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
            Height          =   735
            Left            =   2580
            TabIndex        =   296
            Top             =   180
            Width           =   1665
            Begin MSMask.MaskEdBox txtdata_venc 
               Height          =   315
               Left            =   180
               TabIndex        =   119
               ToolTipText     =   "Data de vencimento."
               Top             =   270
               Width           =   1005
               _ExtentX        =   1773
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
            Begin VB.Image imgCalendario 
               Height          =   360
               Left            =   1170
               Picture         =   "frmcompras_fornecedores.frx":1DCA6
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   240
               Width           =   330
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alíq. ICMS ind."
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
            Index           =   23
            Left            =   7027
            TabIndex        =   309
            Top             =   1500
            Width           =   1050
         End
         Begin VB.Label Label27 
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
            Left            =   285
            TabIndex        =   305
            Top             =   1500
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   10
            Left            =   10830
            TabIndex        =   298
            Top             =   930
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instituição bancária prevista para pagamento"
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
            Index           =   22
            Left            =   2677
            TabIndex        =   297
            Top             =   1500
            Width           =   3270
         End
      End
      Begin VB.TextBox txtIDCliente 
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
         Left            =   225
         MaxLength       =   9
         TabIndex        =   32
         ToolTipText     =   "Códido do fornecedor"
         Top             =   2880
         Width           =   825
      End
      Begin MSComctlLib.ListView ListaEntrega 
         Height          =   6315
         Left            =   -74925
         TabIndex        =   78
         Top             =   3390
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11139
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
            Object.Tag             =   "T"
            Text            =   "Endereço"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Bairro"
            Object.Width           =   7938
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cidade"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "UF"
            Object.Width           =   882
         EndProperty
      End
      Begin VB.TextBox txtid_cobranca 
         Height          =   315
         Left            =   -74280
         TabIndex        =   242
         Text            =   "0"
         Top             =   6420
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtid_entrega 
         Height          =   315
         Left            =   -72900
         TabIndex        =   229
         Text            =   "0"
         Top             =   6030
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   224
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox txtAgencia 
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
            Left            =   10701
            MaxLength       =   50
            TabIndex        =   108
            ToolTipText     =   "Agência."
            Top             =   390
            Width           =   2085
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
            Height          =   315
            Left            =   12810
            MaxLength       =   20
            TabIndex        =   109
            ToolTipText     =   "Conta corrente."
            Top             =   390
            Width           =   2175
         End
         Begin VB.TextBox txtBanco 
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
            Left            =   4034
            MaxLength       =   50
            TabIndex        =   107
            ToolTipText     =   "Banco."
            Top             =   390
            Width           =   6640
         End
         Begin VB.TextBox txtdata5 
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
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox TxtResponsavel5 
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
            Left            =   1387
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2625
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Agência"
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
            Index           =   8
            Left            =   11458
            TabIndex        =   277
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   7
            Left            =   7124
            TabIndex        =   228
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
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
            Left            =   13680
            TabIndex        =   227
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label2 
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
            Index           =   14
            Left            =   600
            TabIndex        =   226
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
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
            Index           =   15
            Left            =   2242
            TabIndex        =   225
            Top             =   180
            Width           =   915
         End
      End
      Begin VB.TextBox txtid_banco 
         Alignment       =   2  'Center
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
         Height          =   335
         Left            =   -72495
         Locked          =   -1  'True
         MouseIcon       =   "frmcompras_fornecedores.frx":1E129
         MousePointer    =   99  'Custom
         TabIndex        =   223
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   4170
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74925
         TabIndex        =   219
         Top             =   1285
         Width           =   15195
         Begin VB.ComboBox cmbFamilia 
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
            Left            =   4410
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   103
            ToolTipText     =   "Família."
            Top             =   390
            Width           =   10605
         End
         Begin VB.TextBox txtdata4 
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
            TabIndex        =   101
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox TxtResponsavel4 
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
            TabIndex        =   102
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3015
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   11
            Left            =   9472
            TabIndex        =   222
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
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
            Index           =   12
            Left            =   600
            TabIndex        =   221
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
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
            Index           =   13
            Left            =   2430
            TabIndex        =   220
            Top             =   180
            Width           =   915
         End
      End
      Begin VB.TextBox txtid_familia 
         Height          =   315
         Left            =   -73365
         TabIndex        =   218
         Text            =   "0"
         Top             =   4380
         Visible         =   0   'False
         Width           =   555
      End
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
         Left            =   -73410
         MaxLength       =   60
         MouseIcon       =   "frmcompras_fornecedores.frx":1E433
         MousePointer    =   99  'Custom
         TabIndex        =   210
         Text            =   "0"
         ToolTipText     =   "Digite o nome para contato."
         Top             =   5400
         Visible         =   0   'False
         Width           =   950
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informações da ordem"
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
         Height          =   3075
         Left            =   -74970
         TabIndex        =   158
         Top             =   1170
         Width           =   11865
         Begin VB.CommandButton cmdcodordem 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   5010
            MouseIcon       =   "frmcompras_fornecedores.frx":1E73D
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":1E88F
            Style           =   1  'Graphical
            TabIndex        =   187
            ToolTipText     =   "Filtrar código da ordem"
            Top             =   407
            Width           =   375
         End
         Begin VB.TextBox txtreferencia 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1110
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":1ECAA
            MousePointer    =   99  'Custom
            TabIndex        =   186
            ToolTipText     =   "Código de referência do item."
            Top             =   1845
            Width           =   1545
         End
         Begin VB.CommandButton cmdreferencia 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   2670
            MouseIcon       =   "frmcompras_fornecedores.frx":1EFB4
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":1F106
            Style           =   1  'Graphical
            TabIndex        =   185
            ToolTipText     =   "Filtrar por numero da referencia do item"
            Top             =   1847
            Width           =   375
         End
         Begin VB.CheckBox chkEspecial 
            Alignment       =   1  'Right Justify
            Caption         =   "Ordem de fabricação sem prioridade"
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
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   2700
            Width           =   11745
         End
         Begin VB.CommandButton cmdordem 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   2670
            MouseIcon       =   "frmcompras_fornecedores.frx":1F521
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":1F673
            Style           =   1  'Graphical
            TabIndex        =   183
            ToolTipText     =   "Filtrar numero da ordem"
            Top             =   407
            Width           =   375
         End
         Begin VB.CommandButton cmdfiltro_item 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   2670
            MouseIcon       =   "frmcompras_fornecedores.frx":1FA8E
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":1FBE0
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "Filtrar por código do item"
            Top             =   1127
            Width           =   375
         End
         Begin VB.TextBox txtop 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   4155
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":1FFFB
            MousePointer    =   99  'Custom
            TabIndex        =   181
            ToolTipText     =   "Código da ordem."
            Top             =   405
            Width           =   810
         End
         Begin VB.TextBox txtlista 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1110
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":20305
            MousePointer    =   99  'Custom
            TabIndex        =   180
            ToolTipText     =   "Número do pedido interno."
            Top             =   758
            Width           =   1125
         End
         Begin VB.CommandButton cmdfiltro_prazo 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   5010
            MouseIcon       =   "frmcompras_fornecedores.frx":2060F
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":20761
            Style           =   1  'Graphical
            TabIndex        =   179
            ToolTipText     =   "Filtrar por prazo final"
            Top             =   767
            Width           =   375
         End
         Begin VB.CommandButton cmdfiltro_nproduto 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   2670
            MouseIcon       =   "frmcompras_fornecedores.frx":20B7C
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":20CCE
            Style           =   1  'Graphical
            TabIndex        =   178
            ToolTipText     =   "Filtrar por código do produto"
            Top             =   1487
            Width           =   375
         End
         Begin VB.CommandButton cmdfiltro_produto 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   11400
            MouseIcon       =   "frmcompras_fornecedores.frx":210E9
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":2123B
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "Filtrar O.F por descrição do produto."
            Top             =   1487
            Width           =   375
         End
         Begin VB.CommandButton cmdfiltro_componente 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   11400
            MouseIcon       =   "frmcompras_fornecedores.frx":21656
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":217A8
            Style           =   1  'Graphical
            TabIndex        =   176
            ToolTipText     =   "Filtrar O.F por descrição do componente."
            Top             =   1127
            Width           =   375
         End
         Begin VB.CommandButton cmdfiltro_cliente 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   11400
            MouseIcon       =   "frmcompras_fornecedores.frx":21BC3
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":21D15
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "Filtrar O.F por cliente."
            Top             =   1847
            Width           =   375
         End
         Begin VB.CommandButton cmdfiltro_pedido 
            Appearance      =   0  'Flat
            Height          =   310
            Left            =   2670
            MouseIcon       =   "frmcompras_fornecedores.frx":22130
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":22282
            Style           =   1  'Graphical
            TabIndex        =   174
            ToolTipText     =   "Filtrar por nº do pedido."
            Top             =   760
            Width           =   375
         End
         Begin VB.TextBox txtcliente 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   4155
            Locked          =   -1  'True
            MaxLength       =   50
            MouseIcon       =   "frmcompras_fornecedores.frx":2269D
            MousePointer    =   99  'Custom
            TabIndex        =   173
            ToolTipText     =   "Nome do cliente."
            Top             =   1845
            Width           =   7200
         End
         Begin VB.TextBox txtnomelista 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   4155
            MaxLength       =   255
            MouseIcon       =   "frmcompras_fornecedores.frx":229A7
            MousePointer    =   99  'Custom
            TabIndex        =   172
            ToolTipText     =   "Descrição do produto."
            Top             =   1485
            Width           =   7200
         End
         Begin VB.TextBox txtmaquina 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1110
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":22CB1
            MousePointer    =   99  'Custom
            TabIndex        =   171
            ToolTipText     =   "Código do produto."
            Top             =   1485
            Width           =   1545
         End
         Begin VB.TextBox txtdesenho 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1110
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":22FBB
            MousePointer    =   99  'Custom
            TabIndex        =   170
            ToolTipText     =   "Código do item."
            Top             =   1125
            Width           =   1545
         End
         Begin VB.TextBox txtProduto 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   4155
            MaxLength       =   255
            MouseIcon       =   "frmcompras_fornecedores.frx":232C5
            MousePointer    =   99  'Custom
            TabIndex        =   169
            ToolTipText     =   "Descrição do item."
            Top             =   1110
            Width           =   7200
         End
         Begin VB.TextBox txtOF 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1110
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":235CF
            MousePointer    =   99  'Custom
            TabIndex        =   168
            ToolTipText     =   "Número da ordem."
            Top             =   405
            Width           =   1545
         End
         Begin VB.Frame Frame15 
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
            Height          =   540
            Left            =   0
            TabIndex        =   164
            Top             =   2160
            Width           =   11865
            Begin VB.CheckBox optMontagem 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Montagem"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   9960
               TabIndex        =   167
               Top             =   180
               Width           =   1335
            End
            Begin VB.CheckBox optExpedicao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Expedição"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   5460
               TabIndex        =   166
               Top             =   180
               Width           =   1335
            End
            Begin VB.CheckBox optFabricacao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fabricação"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   720
               TabIndex        =   165
               Top             =   180
               Width           =   1275
            End
         End
         Begin VB.TextBox txtquantidade 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   10785
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":238D9
            MousePointer    =   99  'Custom
            TabIndex        =   163
            ToolTipText     =   "Quantidade de peças para produzir."
            Top             =   405
            Width           =   990
         End
         Begin VB.TextBox txtnf 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   7455
            MaxLength       =   10
            MouseIcon       =   "frmcompras_fornecedores.frx":23BE3
            MousePointer    =   99  'Custom
            TabIndex        =   162
            ToolTipText     =   "Número da nota fiscal do material."
            Top             =   765
            Width           =   930
         End
         Begin VB.TextBox txtcorrida 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   10905
            MaxLength       =   15
            MouseIcon       =   "frmcompras_fornecedores.frx":23EED
            MousePointer    =   99  'Custom
            TabIndex        =   161
            ToolTipText     =   "Número da corrida do material."
            Top             =   765
            Width           =   870
         End
         Begin VB.TextBox txtconcluida 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   9045
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":241F7
            MousePointer    =   99  'Custom
            TabIndex        =   160
            ToolTipText     =   "Data de conclusão da ordem."
            Top             =   405
            Width           =   990
         End
         Begin VB.TextBox txtrev 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   2220
            MaxLength       =   20
            MouseIcon       =   "frmcompras_fornecedores.frx":24501
            MousePointer    =   99  'Custom
            TabIndex        =   159
            ToolTipText     =   "Número da revisão."
            Top             =   758
            Width           =   435
         End
         Begin MSMask.MaskEdBox mskprazofina 
            Height          =   300
            Left            =   4155
            TabIndex        =   188
            ToolTipText     =   "Data de entrega."
            Top             =   772
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483628
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdtemissao 
            Height          =   300
            Left            =   6135
            TabIndex        =   189
            ToolTipText     =   "Data de emissão da NF."
            Top             =   765
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483628
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Referência :"
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
            Left            =   210
            TabIndex        =   206
            Top             =   1905
            Width           =   885
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. ordem :"
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
            Left            =   3150
            TabIndex        =   205
            Top             =   465
            Width           =   945
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente :"
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
            Left            =   2850
            TabIndex        =   204
            Top             =   1875
            Width           =   1245
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Desc. Prod.:"
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
            Index           =   44
            Left            =   2850
            TabIndex        =   203
            Top             =   1515
            Width           =   1245
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Pedido :"
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
            Left            =   285
            TabIndex        =   202
            Top             =   825
            Width           =   810
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Prod. :"
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
            Left            =   60
            TabIndex        =   201
            Top             =   1545
            Width           =   1035
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo final :"
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
            Left            =   3240
            TabIndex        =   200
            Top             =   825
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Desc. item  :"
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
            Index           =   1
            Left            =   2730
            TabIndex        =   199
            Top             =   1155
            Width           =   1365
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Item :"
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
            Left            =   120
            TabIndex        =   198
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Ordem_ N° :"
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
            Left            =   195
            TabIndex        =   197
            Top             =   465
            Width           =   900
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Certificado:"
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
            Left            =   8430
            TabIndex        =   196
            Top             =   825
            Width           =   840
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Quant.:"
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
            Left            =   10155
            TabIndex        =   195
            Top             =   465
            Width           =   570
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data de Implantação:"
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
            Left            =   5475
            TabIndex        =   194
            Top             =   465
            Width           =   1560
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "N.F.:"
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
            Left            =   7035
            TabIndex        =   193
            Top             =   825
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Corrida:"
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
            Left            =   10290
            TabIndex        =   192
            Top             =   825
            Width           =   585
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Emissão:"
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
            Left            =   5490
            TabIndex        =   191
            Top             =   825
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Conclusão:"
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
            Left            =   8190
            TabIndex        =   190
            Top             =   465
            Width           =   795
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtro da lista de Ordens"
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
         Height          =   780
         Left            =   -74970
         TabIndex        =   151
         Top             =   4260
         Width           =   11865
         Begin VB.CommandButton cmdconcluidas 
            Height          =   375
            Left            =   9480
            MouseIcon       =   "frmcompras_fornecedores.frx":2480B
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":24B15
            Style           =   1  'Graphical
            TabIndex        =   157
            ToolTipText     =   "Imprimir lista de OF's concluidas"
            Top             =   315
            Width           =   375
         End
         Begin VB.CommandButton cmdtodas 
            Height          =   375
            Left            =   4995
            MouseIcon       =   "frmcompras_fornecedores.frx":24C0B
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":24F15
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Imprimir lista de todas OF's emitidas"
            Top             =   315
            Width           =   375
         End
         Begin VB.CommandButton cmdnaoprontas 
            Height          =   375
            Left            =   270
            MouseIcon       =   "frmcompras_fornecedores.frx":2500B
            MousePointer    =   99  'Custom
            Picture         =   "frmcompras_fornecedores.frx":25315
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "Imprimir lista de OF's não concluidas"
            Top             =   300
            Width           =   375
         End
         Begin VB.OptionButton optPronta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Concluidas"
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
            Height          =   252
            Left            =   9960
            TabIndex        =   154
            ToolTipText     =   "Filtra OF`s prontas."
            Top             =   375
            Width           =   1575
         End
         Begin VB.OptionButton opttodas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Emitidas"
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
            Height          =   252
            Left            =   5475
            TabIndex        =   153
            ToolTipText     =   "Filtra todas as OF`s."
            Top             =   375
            Width           =   1335
         End
         Begin VB.OptionButton optnaoprontas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pendentes"
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
            Height          =   252
            Left            =   750
            TabIndex        =   152
            ToolTipText     =   "Filtra OF`s não prontas."
            Top             =   375
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView Lista_contato 
         Height          =   6950
         Left            =   -74925
         TabIndex        =   56
         Top             =   2760
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12250
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Object.Tag             =   "T"
            Text            =   "Nome"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Departamento"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Telefones"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "E-mail"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView lista_familia 
         Height          =   7545
         Left            =   -74925
         TabIndex        =   104
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13309
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   20285
         EndProperty
      End
      Begin MSComctlLib.ListView lista_banco 
         Height          =   7545
         Left            =   -74925
         TabIndex        =   110
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13309
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Banco"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Agência"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Conta"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74925
         TabIndex        =   230
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox txtsite_entrega 
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
            Left            =   10350
            MaxLength       =   60
            TabIndex        =   77
            ToolTipText     =   "Site."
            Top             =   1620
            Width           =   4665
         End
         Begin VB.TextBox txtemail_entrega 
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
            Left            =   5640
            MaxLength       =   60
            TabIndex        =   76
            ToolTipText     =   "E-mail."
            Top             =   1620
            Width           =   4695
         End
         Begin VB.ComboBox cmbTipo_endereco_entrega 
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
            ItemData        =   "frmcompras_fornecedores.frx":2540B
            Left            =   5946
            List            =   "frmcompras_fornecedores.frx":25445
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   60
            ToolTipText     =   "Tipo do endereço."
            Top             =   390
            Width           =   1260
         End
         Begin VB.ComboBox cmbTipo_bairro_entrega 
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
            ItemData        =   "frmcompras_fornecedores.frx":254CD
            Left            =   180
            List            =   "frmcompras_fornecedores.frx":2550D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   64
            ToolTipText     =   "Tipo do bairro."
            Top             =   990
            Width           =   1305
         End
         Begin VB.TextBox txtComplemento_entrega 
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
            Left            =   13890
            MaxLength       =   30
            TabIndex        =   63
            ToolTipText     =   "Complemento."
            Top             =   390
            Width           =   1125
         End
         Begin VB.TextBox txtdata2 
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
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox TxtResponsavel2 
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
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2835
         End
         Begin VB.TextBox txtEndereco_entrega 
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
            Left            =   7200
            MaxLength       =   60
            TabIndex        =   61
            ToolTipText     =   "Endereço."
            Top             =   390
            Width           =   5655
         End
         Begin VB.TextBox txtBairro_entrega 
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
            Left            =   1480
            MaxLength       =   60
            TabIndex        =   65
            ToolTipText     =   "Bairro."
            Top             =   990
            Width           =   3425
         End
         Begin VB.ComboBox cmbuf_entrega 
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
            ItemData        =   "frmcompras_fornecedores.frx":255BD
            Left            =   4920
            List            =   "frmcompras_fornecedores.frx":255BF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   66
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   740
         End
         Begin VB.TextBox txtcxpostal_entrega 
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
            Left            =   9250
            MaxLength       =   50
            TabIndex        =   69
            ToolTipText     =   "Caixa postal."
            Top             =   990
            Width           =   1115
         End
         Begin VB.TextBox mskcep_entrega 
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
            Left            =   10380
            MaxLength       =   12
            TabIndex        =   70
            ToolTipText     =   "CEP."
            Top             =   990
            Width           =   1005
         End
         Begin VB.TextBox txtNumero_entrega 
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
            Left            =   12870
            MaxLength       =   60
            TabIndex        =   62
            ToolTipText     =   "Número."
            Top             =   390
            Width           =   1005
         End
         Begin MSMask.MaskEdBox txttel2_entrega 
            Height          =   315
            Left            =   180
            TabIndex        =   73
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel3_entrega 
            Height          =   315
            Left            =   1995
            TabIndex        =   74
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1820
            _ExtentX        =   3201
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel1_entrega 
            Height          =   315
            Left            =   11400
            TabIndex        =   71
            ToolTipText     =   "Número do telefone."
            Top             =   990
            Width           =   1785
            _ExtentX        =   3149
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel4_entrega 
            Height          =   315
            Left            =   3825
            TabIndex        =   75
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfax_entrega 
            Height          =   315
            Left            =   13200
            TabIndex        =   72
            ToolTipText     =   "Número do fax."
            Top             =   990
            Width           =   1815
            _ExtentX        =   3201
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCNPJ_entrega 
            Height          =   315
            Left            =   4230
            TabIndex        =   59
            ToolTipText     =   "Número do CNPJ."
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbCidade_Entrega 
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
            ItemData        =   "frmcompras_fornecedores.frx":255C1
            Left            =   5670
            List            =   "frmcompras_fornecedores.frx":255C3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   67
            ToolTipText     =   "Cidade."
            Top             =   990
            Width           =   3570
         End
         Begin VB.TextBox txtCidade_entrega 
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
            Left            =   5670
            MaxLength       =   60
            TabIndex        =   68
            ToolTipText     =   "Cidade."
            Top             =   990
            Visible         =   0   'False
            Width           =   3570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
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
            Index           =   56
            Left            =   12540
            TabIndex        =   293
            Top             =   1410
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.04"
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
            Index           =   31
            Left            =   4485
            TabIndex        =   292
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.03"
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
            Index           =   30
            Left            =   2685
            TabIndex        =   291
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.02"
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
            Left            =   855
            TabIndex        =   290
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   47
            Left            =   6426
            TabIndex        =   270
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   48
            Left            =   682
            TabIndex        =   269
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            Index           =   51
            Left            =   13965
            TabIndex        =   268
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label2 
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
            Index           =   16
            Left            =   600
            TabIndex        =   262
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
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
            Index           =   7
            Left            =   2340
            TabIndex        =   261
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cx. postal"
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
            Index           =   20
            Left            =   9435
            TabIndex        =   241
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Index           =   21
            Left            =   10740
            TabIndex        =   240
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Index           =   12
            Left            =   13980
            TabIndex        =   239
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
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
            Index           =   22
            Left            =   7770
            TabIndex        =   238
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
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
            Index           =   24
            Left            =   9690
            TabIndex        =   237
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
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
            Index           =   25
            Left            =   2982
            TabIndex        =   236
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   26
            Left            =   7215
            TabIndex        =   235
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
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
            Index           =   27
            Left            =   5200
            TabIndex        =   234
            Top             =   780
            Width           =   180
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.01"
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
            Index           =   28
            Left            =   12075
            TabIndex        =   233
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ"
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
            Left            =   4875
            TabIndex        =   232
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   13095
            TabIndex        =   231
            Top             =   180
            Width           =   555
         End
      End
      Begin MSComctlLib.ListView listacobranca 
         Height          =   6315
         Left            =   -74925
         TabIndex        =   100
         Top             =   3390
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11139
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
            Object.Tag             =   "T"
            Text            =   "Endereço"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Bairro"
            Object.Width           =   7938
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cidade"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "UF"
            Object.Width           =   882
         EndProperty
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74925
         TabIndex        =   243
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox txtSite_cobranca 
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
            Left            =   10350
            MaxLength       =   60
            TabIndex        =   99
            ToolTipText     =   "Site."
            Top             =   1620
            Width           =   4665
         End
         Begin VB.TextBox txtemail_cobranca 
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
            Left            =   5640
            MaxLength       =   60
            TabIndex        =   98
            ToolTipText     =   "E-mail."
            Top             =   1620
            Width           =   4695
         End
         Begin VB.ComboBox cmbTipo_endereco_cobranca 
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
            ItemData        =   "frmcompras_fornecedores.frx":255C5
            Left            =   5940
            List            =   "frmcompras_fornecedores.frx":255FF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   82
            ToolTipText     =   "Tipo do endereço."
            Top             =   390
            Width           =   1260
         End
         Begin VB.ComboBox cmbTipo_bairro_cobranca 
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
            ItemData        =   "frmcompras_fornecedores.frx":25687
            Left            =   180
            List            =   "frmcompras_fornecedores.frx":256C7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   86
            ToolTipText     =   "Tipo do bairro."
            Top             =   990
            Width           =   1305
         End
         Begin VB.TextBox txtComplemento_cobranca 
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
            Left            =   13890
            MaxLength       =   30
            TabIndex        =   85
            ToolTipText     =   "Complemento."
            Top             =   390
            Width           =   1125
         End
         Begin VB.TextBox txtdata3 
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
            TabIndex        =   79
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox TxtResponsavel3 
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
            TabIndex        =   80
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2835
         End
         Begin VB.TextBox txtcxpostal_cobranca 
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
            Left            =   9250
            MaxLength       =   50
            TabIndex        =   91
            ToolTipText     =   "Caixa postal."
            Top             =   990
            Width           =   1115
         End
         Begin VB.ComboBox cmbuf_cobranca 
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
            ItemData        =   "frmcompras_fornecedores.frx":25777
            Left            =   4920
            List            =   "frmcompras_fornecedores.frx":25779
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   88
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   740
         End
         Begin VB.TextBox txtbairro_cobranca 
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
            Left            =   1480
            MaxLength       =   60
            TabIndex        =   87
            ToolTipText     =   "Bairro."
            Top             =   990
            Width           =   3425
         End
         Begin VB.TextBox txtendereco_cobranca 
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
            Left            =   7200
            MaxLength       =   60
            TabIndex        =   83
            ToolTipText     =   "Endereço."
            Top             =   390
            Width           =   5655
         End
         Begin VB.TextBox mskcep_cobranca 
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
            Left            =   10380
            MaxLength       =   12
            TabIndex        =   92
            ToolTipText     =   "CEP."
            Top             =   990
            Width           =   1005
         End
         Begin VB.TextBox txtNumero_cobranca 
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
            Left            =   12870
            MaxLength       =   60
            TabIndex        =   84
            ToolTipText     =   "Número."
            Top             =   390
            Width           =   1005
         End
         Begin MSMask.MaskEdBox txttel2_cobranca 
            Height          =   315
            Left            =   180
            TabIndex        =   95
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel3_cobranca 
            Height          =   315
            Left            =   2010
            TabIndex        =   96
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel1_cobranca 
            Height          =   315
            Left            =   11400
            TabIndex        =   93
            ToolTipText     =   "Número do telefone."
            Top             =   990
            Width           =   1785
            _ExtentX        =   3149
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttel4_cobranca 
            Height          =   315
            Left            =   3825
            TabIndex        =   97
            ToolTipText     =   "Número do telefone."
            Top             =   1620
            Width           =   1800
            _ExtentX        =   3175
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfax_cobranca 
            Height          =   315
            Left            =   13200
            TabIndex        =   94
            ToolTipText     =   "Número do fax."
            Top             =   990
            Width           =   1815
            _ExtentX        =   3201
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCNPJ_cobranca 
            Height          =   315
            Left            =   4230
            TabIndex        =   81
            ToolTipText     =   "Número do CNPJ."
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbCidade_cobranca 
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
            ItemData        =   "frmcompras_fornecedores.frx":2577B
            Left            =   5670
            List            =   "frmcompras_fornecedores.frx":2577D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   89
            ToolTipText     =   "Cidade."
            Top             =   990
            Width           =   3570
         End
         Begin VB.TextBox txtcidade_cobranca 
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
            Left            =   5670
            MaxLength       =   60
            TabIndex        =   90
            ToolTipText     =   "Cidade."
            Top             =   990
            Visible         =   0   'False
            Width           =   3570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
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
            Index           =   57
            Left            =   12540
            TabIndex        =   294
            Top             =   1410
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   49
            Left            =   6420
            TabIndex        =   273
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   50
            Left            =   682
            TabIndex        =   272
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            Index           =   52
            Left            =   13965
            TabIndex        =   271
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label2 
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
            Index           =   18
            Left            =   600
            TabIndex        =   264
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
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
            Index           =   17
            Left            =   2340
            TabIndex        =   263
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.04"
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
            Index           =   18
            Left            =   4485
            TabIndex        =   257
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.03"
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
            Index           =   19
            Left            =   2685
            TabIndex        =   256
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.02"
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
            Index           =   23
            Left            =   855
            TabIndex        =   255
            Top             =   1410
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.01"
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
            Index           =   32
            Left            =   12075
            TabIndex        =   254
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
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
            Height          =   255
            Index           =   33
            Left            =   5185
            TabIndex        =   253
            Top             =   780
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   34
            Left            =   7215
            TabIndex        =   252
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
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
            Index           =   35
            Left            =   2982
            TabIndex        =   251
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
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
            Index           =   36
            Left            =   9690
            TabIndex        =   250
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
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
            Index           =   37
            Left            =   7770
            TabIndex        =   249
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Left            =   13980
            TabIndex        =   248
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Index           =   38
            Left            =   10740
            TabIndex        =   247
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cx. postal"
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
            Index           =   39
            Left            =   9435
            TabIndex        =   246
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ"
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
            Left            =   4875
            TabIndex        =   245
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   13095
            TabIndex        =   244
            Top             =   180
            Width           =   555
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74940
         TabIndex        =   299
         Top             =   330
         Width           =   15195
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
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   12930
            Top             =   120
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmcompras_fornecedores.frx":2577F
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_Segmento 
         Height          =   7545
         Left            =   -74925
         TabIndex        =   115
         Top             =   2160
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13309
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Segmento"
            Object.Width           =   20285
         EndProperty
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74925
         TabIndex        =   133
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox TxtEmail_Contato 
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
            MaxLength       =   60
            TabIndex        =   53
            ToolTipText     =   "E-mail."
            Top             =   990
            Width           =   6285
         End
         Begin VB.CheckBox Chk_enviar_boleto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Enviar boleto"
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
            Left            =   13770
            TabIndex        =   55
            Top             =   1110
            Width           =   1245
         End
         Begin VB.CheckBox Chk_enviar_NFe 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Enviar NFe"
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
            Left            =   13770
            TabIndex        =   54
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox TxtResponsavel1 
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
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   5175
         End
         Begin VB.TextBox txtdata1 
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
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1185
         End
         Begin VB.TextBox txttelcontato 
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
            MaxLength       =   20
            TabIndex        =   52
            ToolTipText     =   "Telefones."
            Top             =   990
            Width           =   3375
         End
         Begin VB.TextBox txtdepartamento 
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
            TabIndex        =   51
            ToolTipText     =   "Departamento."
            Top             =   990
            Width           =   3735
         End
         Begin VB.TextBox txtnomecontato 
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
            MaxLength       =   100
            TabIndex        =   50
            ToolTipText     =   "Nome."
            Top             =   390
            Width           =   8445
         End
         Begin VB.Label Label2 
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
            Left            =   3285
            TabIndex        =   215
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Left            =   600
            TabIndex        =   214
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
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
            Left            =   10252
            TabIndex        =   138
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefones"
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
            Index           =   14
            Left            =   5265
            TabIndex        =   137
            Top             =   780
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
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
            Left            =   10200
            TabIndex        =   136
            Top             =   180
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            Left            =   1530
            TabIndex        =   135
            Top             =   780
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   75
         TabIndex        =   139
         Top             =   1285
         Width           =   15195
         Begin VB.TextBox txtSite 
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
            Left            =   6240
            MaxLength       =   100
            TabIndex        =   24
            ToolTipText     =   "Site."
            Top             =   2775
            Width           =   2775
         End
         Begin VB.TextBox txtemail 
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
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   23
            ToolTipText     =   "E-mail."
            Top             =   2775
            Width           =   2745
         End
         Begin VB.CheckBox Chk_nao_contribuinte_ICMS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Não contribuinte ICMS"
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
            Left            =   13110
            TabIndex        =   38
            Top             =   510
            Width           =   1935
         End
         Begin VB.CheckBox Chk_enviar_NF 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Enviar NF"
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
            Left            =   11880
            TabIndex        =   34
            Top             =   270
            Width           =   1095
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Data de validação."
            Top             =   390
            Width           =   2025
         End
         Begin VB.ComboBox cmbRegimeTributario 
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
            ItemData        =   "frmcompras_fornecedores.frx":29503
            Left            =   7980
            List            =   "frmcompras_fornecedores.frx":29505
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Regime tributário."
            Top             =   990
            Width           =   1635
         End
         Begin VB.ComboBox cmbPessoa 
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
            ItemData        =   "frmcompras_fornecedores.frx":29507
            Left            =   165
            List            =   "frmcompras_fornecedores.frx":29511
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Tipo."
            Top             =   990
            Width           =   1170
         End
         Begin VB.CheckBox Chk_prospecto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prospecto"
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
            Left            =   11880
            TabIndex        =   33
            Top             =   510
            Width           =   1035
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
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   3735
         End
         Begin VB.ComboBox Cmb_tipo_transp 
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
            ItemData        =   "frmcompras_fornecedores.frx":29527
            Left            =   9030
            List            =   "frmcompras_fornecedores.frx":29537
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Tipo da transportadora."
            Top             =   2775
            Width           =   1215
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
            Left            =   8400
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   37
            ToolTipText     =   "Centro de custo."
            Top             =   3480
            Width           =   4395
         End
         Begin VB.ComboBox cmbTipo_bairro 
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
            ItemData        =   "frmcompras_fornecedores.frx":2955B
            Left            =   3870
            List            =   "frmcompras_fornecedores.frx":2959B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Tipo do bairro."
            Top             =   2190
            Width           =   1125
         End
         Begin VB.ComboBox cmbTipo_endereco 
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
            ItemData        =   "frmcompras_fornecedores.frx":2964B
            Left            =   8190
            List            =   "frmcompras_fornecedores.frx":2968B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Tipo do endereço."
            Top             =   2190
            Width           =   1050
         End
         Begin VB.TextBox txtComplemento 
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
            Height          =   325
            Left            =   14040
            MaxLength       =   30
            TabIndex        =   19
            ToolTipText     =   "Complemento."
            Top             =   2190
            Width           =   1035
         End
         Begin VB.ComboBox Txt_pais 
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
            ItemData        =   "frmcompras_fornecedores.frx":29722
            Left            =   2550
            List            =   "frmcompras_fornecedores.frx":29724
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "País."
            Top             =   990
            Width           =   2235
         End
         Begin VB.TextBox txtcep 
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
            MaxLength       =   12
            TabIndex        =   12
            ToolTipText     =   "CEP."
            Top             =   2190
            Width           =   885
         End
         Begin VB.TextBox txtcaixapostal 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Caixa postal."
            Top             =   2775
            Width           =   795
         End
         Begin VB.ComboBox cmbtransportadora 
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
            Left            =   10260
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Nome da transportadora."
            Top             =   2775
            Width           =   4455
         End
         Begin VB.TextBox txtNumero 
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
            Height          =   325
            Left            =   13455
            MaxLength       =   60
            TabIndex        =   18
            ToolTipText     =   "Número."
            Top             =   2190
            Width           =   570
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
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4575
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
            Left            =   11970
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   990
            Width           =   2475
         End
         Begin VB.TextBox txtIM_IE 
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
            Left            =   10860
            MaxLength       =   15
            TabIndex        =   7
            ToolTipText     =   "Inscrição municipal."
            Top             =   990
            Width           =   1095
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
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   975
         End
         Begin VB.ComboBox TXTcategoria 
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
            ItemData        =   "frmcompras_fornecedores.frx":29726
            Left            =   14460
            List            =   "frmcompras_fornecedores.frx":29728
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "IQF."
            Top             =   990
            Width           =   615
         End
         Begin VB.TextBox txtnomefantasia 
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
            Left            =   7440
            MaxLength       =   60
            TabIndex        =   11
            ToolTipText     =   "Nome fantasia."
            Top             =   1590
            Width           =   7605
         End
         Begin VB.TextBox txtNomeRazao 
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
            Left            =   990
            MaxLength       =   60
            TabIndex        =   10
            ToolTipText     =   "Razão social."
            Top             =   1590
            Width           =   6435
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
            Height          =   325
            Left            =   9255
            MaxLength       =   60
            TabIndex        =   17
            ToolTipText     =   "Endereço."
            Top             =   2190
            Width           =   4185
         End
         Begin VB.TextBox txtBairro 
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
            Height          =   325
            Left            =   5010
            MaxLength       =   60
            TabIndex        =   15
            ToolTipText     =   "Bairro."
            Top             =   2190
            Width           =   3180
         End
         Begin MSMask.MaskEdBox txtTelefones 
            Height          =   330
            Left            =   930
            TabIndex        =   21
            ToolTipText     =   "Número do telefone."
            Top             =   2775
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   582
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtRG_IE 
            Height          =   315
            Left            =   9630
            TabIndex        =   6
            ToolTipText     =   "Inscrição estadual."
            Top             =   990
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFax 
            Height          =   330
            Left            =   2370
            TabIndex        =   22
            ToolTipText     =   "Número do fax."
            Top             =   2775
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCNPJ 
            Height          =   315
            Left            =   5400
            TabIndex        =   4
            ToolTipText     =   "Número do CNPJ."
            Top             =   990
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.###/####-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbUF 
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
            ItemData        =   "frmcompras_fornecedores.frx":2972A
            Left            =   4785
            List            =   "frmcompras_fornecedores.frx":2972C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "UF."
            Top             =   990
            Width           =   615
         End
         Begin MSMask.MaskEdBox txtCpf 
            Height          =   315
            Left            =   5400
            TabIndex        =   35
            ToolTipText     =   "Número do CPF."
            Top             =   990
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbCidade 
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
            ItemData        =   "frmcompras_fornecedores.frx":2972E
            Left            =   1395
            List            =   "frmcompras_fornecedores.frx":29730
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Cidade."
            Top             =   2190
            Width           =   2490
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1410
            MaxLength       =   60
            TabIndex        =   36
            ToolTipText     =   "Cidade."
            Top             =   2190
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.ComboBox cmbOrigem 
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
            ItemData        =   "frmcompras_fornecedores.frx":29732
            Left            =   1350
            List            =   "frmcompras_fornecedores.frx":2973F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Origem."
            Top             =   990
            Width           =   1185
         End
         Begin DrawSuite2022.USButton cmdConsultar 
            Height          =   315
            Left            =   6990
            TabIndex        =   331
            ToolTipText     =   "Consultar cadastro na receita federal."
            Top             =   990
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmcompras_fornecedores.frx":2975C
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   0
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton Cmd_buscarCEP 
            Height          =   315
            Left            =   1050
            TabIndex        =   332
            ToolTipText     =   "Consultar endereço por CEP nos correios"
            Top             =   2190
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmcompras_fornecedores.frx":308EF
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   0
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton cmdLocTransp 
            Height          =   315
            Left            =   14730
            TabIndex        =   27
            ToolTipText     =   "Consultar cadastro transportadora"
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            DibPicture      =   "frmcompras_fornecedores.frx":37A82
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            Theme           =   5
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton btnSintegra 
            Height          =   315
            Left            =   7320
            TabIndex        =   333
            ToolTipText     =   "Consultar cadastro no Sintegra."
            Top             =   990
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmcompras_fornecedores.frx":3D4E7
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   0
            ShowFocusRect   =   0   'False
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton btnRF 
            Height          =   315
            Left            =   7650
            TabIndex        =   334
            ToolTipText     =   "Consultar cadastro na receita federal."
            Top             =   990
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmcompras_fornecedores.frx":45D17
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
            BorderColor     =   8421504
            BorderColorDisabled=   13160660
            BorderColorDown =   7907521
            BorderColorOver =   7907521
            GradientColor2  =   14737632
            GradientColor3  =   12632256
            GradientColor4  =   12632256
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   14417407
            GradientColorOver2=   12317439
            GradientColorOver3=   4838399
            GradientColorOver4=   9627391
            GradientColorDown1=   10802943
            GradientColorDown2=   7979263
            GradientColorDown3=   4370174
            GradientColorDown4=   7395582
            GradientColors  =   1
            PicAlign        =   0
            ShowFocusRect   =   0   'False
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.Label Label26 
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
            Left            =   5925
            TabIndex        =   289
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Regime tributário*"
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
            Height          =   195
            Index           =   54
            Left            =   8100
            TabIndex        =   287
            Top             =   780
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Site"
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
            Index           =   53
            Left            =   7485
            TabIndex        =   286
            Top             =   2580
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origem*"
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
            Height          =   195
            Index           =   55
            Left            =   1680
            TabIndex        =   285
            Top             =   780
            Width           =   600
         End
         Begin VB.Label Label50 
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
            Left            =   8670
            TabIndex        =   284
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   555
            TabIndex        =   283
            Top             =   780
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   5
            Left            =   9420
            TabIndex        =   281
            Top             =   2580
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro de custo"
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
            Index           =   4
            Left            =   10020
            TabIndex        =   274
            Top             =   3270
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   41
            Left            =   4215
            TabIndex        =   267
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
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
            Index           =   40
            Left            =   14085
            TabIndex        =   266
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Index           =   15
            Left            =   8520
            TabIndex        =   265
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
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
            Index           =   12
            Left            =   495
            TabIndex        =   260
            Top             =   1980
            Width           =   285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cx. postal"
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
            Index           =   8
            Left            =   150
            TabIndex        =   259
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transportadora*"
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
            Height          =   195
            Index           =   0
            Left            =   11940
            TabIndex        =   258
            Top             =   2580
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N°"
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
            Left            =   13680
            TabIndex        =   217
            Top             =   1980
            Width           =   180
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "País*"
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
            Height          =   195
            Index           =   42
            Left            =   3525
            TabIndex        =   216
            Top             =   780
            Width           =   375
         End
         Begin VB.Label Label2 
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
            Left            =   3015
            TabIndex        =   213
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
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
            Index           =   3
            Left            =   12975
            TabIndex        =   209
            Top             =   780
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "IM"
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
            Index           =   6
            Left            =   11370
            TabIndex        =   208
            Top             =   780
            Width           =   180
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone (Fax)"
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
            Index           =   6
            Left            =   2385
            TabIndex        =   207
            Top             =   2580
            Width           =   1065
         End
         Begin VB.Label Label2 
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
            Index           =   4
            Left            =   495
            TabIndex        =   150
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome fantasia"
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
            Left            =   10665
            TabIndex        =   149
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IQF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   14625
            TabIndex        =   148
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone"
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
            Left            =   1305
            TabIndex        =   147
            Top             =   2580
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
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
            Index           =   7
            Left            =   4650
            TabIndex        =   146
            Top             =   2580
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IE*"
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
            Height          =   195
            Index           =   2
            Left            =   10110
            TabIndex        =   145
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razão social (60 caracteres máximo NFe)*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   2430
            TabIndex        =   144
            Top             =   1380
            Width           =   3600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
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
            Left            =   11145
            TabIndex        =   143
            Top             =   1980
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro"
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
            Left            =   6375
            TabIndex        =   142
            Top             =   1980
            Width           =   420
         End
         Begin VB.Label Label3 
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
            Index           =   4
            Left            =   2393
            TabIndex        =   141
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF*"
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
            Height          =   195
            Index           =   5
            Left            =   4995
            TabIndex        =   140
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5985
            TabIndex        =   211
            Top             =   780
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   5910
            TabIndex        =   212
            Top             =   780
            Width           =   510
         End
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
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmCompras_fornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Fornecedor   As Boolean 'OK
Dim Novo_Fornecedor1     As Boolean 'OK
Dim Novo_Fornecedor2     As Boolean 'OK
Dim Novo_Fornecedor3     As Boolean 'OK
Dim Novo_Fornecedor4     As Boolean 'OK
Dim Novo_Fornecedor5     As Boolean 'OK
Dim Novo_Fornecedor6     As Boolean 'OK
Public StrSql_Fornecedor As String 'OK
Public FormulaRel_Fornecedor As String 'OK
Dim TBLISTA_Fornecedor   As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

'FunAbrirVideoWeb ("http://www.youtube.com/watch?v=rQxAohwaWxQ&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=57&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Or SSTab1.Tab = 7 Then
    With Lista
        .Visible = True
        If SSTab1.Tab = 0 Then
            .Top = Frame1.Top + Frame1.Height
            .Height = Frame2.Top - .Top
        Else
            .Top = Frame18.Top + Frame18.Height
            .Height = Frame2.Top - .Top
        End If
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRF_Click()
On Error GoTo tratar_erro

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
Dim Plugin As String

Plugin = "RF"

CnpjDados = ReturnNumbersOnly(txtcnpj.Text)

obj.Open "GET", "https://www.sintegraws.com.br/api/v1/execute-api.php?token=1F718E4E-3222-42F1-95D6-995FC9E69C9C&cnpj=" & CnpjDados & "&plugin=" & Plugin & ""
                                                                            
conteudo = CnpjDados
obj.send conteudo

resposta = obj.responseText
'Debug.print resposta

If LerDadosJSON(resposta, "status", "", "") = "OK" And LerDadosJSON(resposta, "code", "", "") = "0" Then

USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = UCase(LerDadosJSON(resposta, "nome", "", ""))
cmbuf.Text = UCase(LerDadosJSON(resposta, "uf", "", ""))
txttel01 = LerDadosJSON(resposta, "telefone", "", "")
txtBairro = UCase(LerDadosJSON(resposta, "bairro", "", ""))
txtendereco = UCase(LerDadosJSON(resposta, "logradouro", "", ""))
txtNumero = LerDadosJSON(resposta, "numero", "", "")
txtCEP = LerDadosJSON(resposta, "cep", "", "")
txtEmail = LerDadosJSON(resposta, "email", "", "")
'cmbCidade.Text = UCase(LerDadosJSON(resposta, "municipio", "", ""))
txtnomefantasia = UCase(LerDadosJSON(resposta, "fantasia", "", ""))
'cmbRegimeTributario.Text = IIf(LerDadosJSON(resposta, "regime_tributacao", "", "") = "Normal - regime periódico de apuração", "Lucro presumido", "Simples Nacional")
'txtRG_IE = Trim(LerDadosJSON(resposta, "inscricao_estadual", "", ""))

txtCategoria.Text = "A"

Cmd_buscarCEP_Click
Else
USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = ""
cmbuf.ListIndex = -1
txtBairro = ""
txtendereco = ""
txtNumero = ""
txtCEP = ""
txtnomefantasia = ""
cmbRegimeTributario.ListIndex = -1
txtRG_IE = ""
txtCategoria.ListIndex = -1

End If


Exit Sub
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("Não foi possível carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnSintegra_Click()
On Error GoTo tratar_erro

Dim resposta As String
Dim obj As MSXML2.ServerXMLHTTP50
Set obj = New MSXML2.ServerXMLHTTP50
Dim Plugin As String

Plugin = "ST"


CnpjDados = ReturnNumbersOnly(txtcnpj.Text)

obj.Open "GET", "https://www.sintegraws.com.br/api/v1/execute-api.php?token=1F718E4E-3222-42F1-95D6-995FC9E69C9C&cnpj=" & CnpjDados & "&plugin=" & Plugin & ""
'1F718E4E-3222-42F1-95D6-995FC9E69C9C'
conteudo = CnpjDados
obj.send conteudo

resposta = obj.responseText
'Debug.print resposta

If LerDadosJSON(resposta, "status", "", "") = "OK" And LerDadosJSON(resposta, "code", "", "") = "0" Then

USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = UCase(LerDadosJSON(resposta, "nome_empresarial", "", ""))
cmbuf.Text = UCase(LerDadosJSON(resposta, "uf", "", ""))
'txttel01 = LerDadosJSON(resposta, "telefone", "", "")
txtBairro = UCase(LerDadosJSON(resposta, "bairro", "", ""))
txtendereco = UCase(LerDadosJSON(resposta, "logradouro", "", ""))
txtNumero = LerDadosJSON(resposta, "numero", "", "")
txtCEP = LerDadosJSON(resposta, "cep", "", "")
'cmbCidade.Text = UCase(LerDadosJSON(resposta, "municipio", "", ""))
txtnomefantasia = UCase(LerDadosJSON(resposta, "nome_fantasia", "", ""))
cmbRegimeTributario.Text = IIf(LerDadosJSON(resposta, "regime_tributacao", "", "") = "Normal - regime periódico de apuração", "Lucro presumido", "Simples Nacional")
txtRG_IE = Trim(LerDadosJSON(resposta, "inscricao_estadual", "", ""))

txtCategoria.Text = "A"

Cmd_buscarCEP_Click
Else
USMsgBox LerDadosJSON(resposta, "message", "", ""), vbInformation, "CAPRIND v5.0"
txtnomerazao.Text = ""
cmbuf.ListIndex = -1
txtBairro = ""
txtendereco = ""
txtNumero = ""
txtCEP = ""
txtnomefantasia = ""
cmbRegimeTributario.ListIndex = -1
txtRG_IE = ""
txtCategoria.ListIndex = -1

End If




Exit Sub
tratar_erro:
    MousePointer = 0
    If Err.Number = 91 Then
        USMsgBox ("Não foi possível carregar todos os dados referentes a este CEP."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub chkAprovada_Cliente_Click()
On Error GoTo tratar_erro

If chkAprovada_Cliente.Value = 1 Then
    Chk_certificado.Value = 0
    Chk_avaliado.Value = 0
    Frame7.Enabled = True
    Frame8.Enabled = True
Else
    If Chk_certificado.Value = 0 And Chk_avaliado.Value = 0 Then
        Frame7.Enabled = False
        Frame8.Enabled = False
        txtdata_venc.Text = "__/__/____"
        txtCertificado = ""
        Txt_caminho_certificado = ""
    End If
End If

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
    If Cmb_opcao_lista = "Validação" Then
        .ButtonState(4) = 5
        .ButtonState(9) = 5
        .ButtonState(10) = 0
    ElseIf Cmb_opcao_lista = "Excluir" Then
            .ButtonState(4) = 0
            .ButtonState(9) = 5
            .ButtonState(10) = 5
        Else
            .ButtonState(4) = 5
            .ButtonState(9) = 0
            .ButtonState(10) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_transp_Click()
On Error GoTo tratar_erro

ProcCarregaComboTransp

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTransp()
On Error GoTo tratar_erro

With cmbtransportadora
    .Clear
    If Cmb_tipo_transp <> "" Then
        If Cmb_tipo_transp = "Empresa" Then
            NomeTabela = "Empresa"
            NomeCampo = "Empresa"
            NomeCampo1 = "Codigo"
        Else
            NomeCampo1 = "IDCliente"
            If Cmb_tipo_transp = "Cliente" Then
                NomeTabela = "Clientes"
                NomeCampo = "NomeRazao"
            Else
                NomeTabela = "Compras_fornecedores"
                NomeCampo = "Nome_Razao"
            End If
        End If
        
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select " & NomeCampo & ", " & NomeCampo1 & " FROM " & NomeTabela & " where " & NomeCampo & " is not null group by " & NomeCampo & ", " & NomeCampo1, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            .AddItem ""
            Do While TBLISTA.EOF = False
                Select Case Cmb_tipo_transp
                    Case "Cliente":
                        .AddItem TBLISTA!NomeRazao
                        .ItemData(.NewIndex) = TBLISTA!IDCliente
                    Case "Fornecedor":
                        .AddItem TBLISTA!Nome_Razao
                        .ItemData(.NewIndex) = TBLISTA!IDCliente
                    Case "Empresa":
                        .AddItem TBLISTA!Empresa
                        .ItemData(.NewIndex) = TBLISTA!CODIGO
                End Select
                TBLISTA.MoveNext
            Loop
        End If
        TBLISTA.Close
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbOrigem_Click()
On Error GoTo tratar_erro

ProcCarregaComboUF cmbuf, "UF is not null", cmbOrigem
ProcCarregaComboUF cmbuf_entrega, "UF is not null", cmbOrigem
ProcCarregaComboUF cmbuf_cobranca, "UF is not null", cmbOrigem
If cmbOrigem = "Estrangeiro" Then
    With txtcnpj
        .Text = "__.___.___/____-__"
        .Enabled = False
    End With
    With txtCpf
        .Text = "___.___.___-__"
        .Enabled = False
    End With
    cmbCidade.Visible = False
    cmbCidade_Entrega.Visible = False
    cmbCidade_cobranca.Visible = False
    txtCidade.Visible = True
    txtCidade_Entrega.Visible = True
    txtcidade_cobranca.Visible = True
Else
    Txt_pais.Text = "BRASIL"
    txtcnpj.Enabled = True
    txtCpf.Enabled = True
    cmbCidade.Visible = True
    cmbCidade_Entrega.Visible = True
    cmbCidade_cobranca.Visible = True
    txtCidade.Visible = False
    txtCidade_Entrega.Visible = False
    txtcidade_cobranca.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPessoa_Click()
On Error GoTo tratar_erro

ProcVerifPessoa

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifPessoa()
On Error GoTo tratar_erro

If cmbPessoa = "FÍSICA" Then
    With txtCpf
        .Text = "___.___.___-__"
        .Enabled = True
        .Visible = True
    End With
    With txtcnpj
        .Text = "__.___.___/____-__"
        .Enabled = False
        .Visible = False
    End With
    With txtIM_IE
        .Text = ""
        .ToolTipText = "Inscrição estadual."
    End With
    Label2(6).Caption = "IE"
    With txtRG_IE
        .Text = ""
        .ToolTipText = "Registro geral."
    End With
    Label16.Visible = True
    Label18.Visible = False
    Label2(2).Caption = "RG"
    
    'Entrega
    Label21(0).Caption = "CPF"
    With txtCNPJ_entrega
        .Mask = "###.###.###-##"
        .ToolTipText = "Número do CPF."
    End With
    
    'Cobrança
    Label24.Caption = "CPF"
    With txtCNPJ_cobranca
        .Mask = "###.###.###-##"
        .ToolTipText = "Número do CPF."
    End With
    
    'Regime
    With cmbRegimeTributario
        .Clear
        .AddItem ""
        .AddItem "MEI"
    End With
Else
    With txtCpf
        .Text = "___.___.___-__"
        .Enabled = False
        .Visible = False
    End With
    With txtcnpj
        .Text = "__.___.___/____-__"
        .Enabled = True
        .Visible = True
    End With
    With txtIM_IE
        .Text = ""
        .ToolTipText = "Inscrição municipal."
    End With
    Label2(6).Caption = "IM"
    With txtRG_IE
        .Text = ""
        .ToolTipText = "Inscrição estadual."
    End With
    Label16.Visible = False
    Label18.Visible = True
    Label2(2).Caption = "IE"
    
    'Entrega
    Label21(0).Caption = "CNPJ"
    With txtCNPJ_entrega
        .Mask = "##.###.###/####-##"
        .ToolTipText = "Número do CNPJ."
    End With
    
    'Cobrança
    Label24.Caption = "CNPJ"
    With txtCNPJ_cobranca
        .Mask = "##.###.###/####-##"
        .ToolTipText = "Número do CNPJ."
    End With
    
    'Regime
    With cmbRegimeTributario
        .Clear
        .AddItem ""
        .AddItem "Lucro presumido"
        .AddItem "Lucro real"
        .AddItem "Simples nacional"
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrar_todos()
On Error GoTo tratar_erro

StrSql_Fornecedor = "Select * from Compras_fornecedores order by nome_razao"
FormulaRel_Fornecedor = ""
ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDcliente = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores order by Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDCliente = " & txtIDcliente)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtIDcliente = TBLISTA!IDCliente
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from Compras_fornecedores where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        procLimpacamposContatos
        ProcLimpacamposEntrega
        ProcLimpacamposCobranca
        ProcLimpaFamilia
        ProcLimpacampos_banco
        ProcLimpaCampos_Outros
        ProcLimpaCampos_Segmento
        ProcPuxaDados
        ProcCarregaListaContatos
        Proccarregalistaentrega
        ProcCarregalistacobranca
        ProcCarregaListaFamilia
        ProcCarregalista_banco
        ProcCarregaListaSegmento
        ProcPuxaDados_Outros
    Else
        USMsgBox ("Fim dos cadastros de fornecedor."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Fornecedor1 = False
Novo_Fornecedor4 = False
Novo_Fornecedor5 = False
Novo_Fornecedor6 = False

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
    USMsgBox ("Informe o(s) fornecedor(es) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Compras_Fornecedores = True
Financeiro_Instituicao = False
frmCompras_fornecedores_bloq.Show 1

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
                If USMsgBox("Deseja realmente excluir este(s) fornecedor(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Compras_fornecedores WHERE idCliente = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from contatos_fornecedor where IdFornecedor = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from compras_fornecedores_familia where IDCliente = " & .ListItems(InitFor) & ""
            Conexao.Execute "DELETE from Compras_fornecedores_banco where id_fornecedor = " & .ListItems(InitFor) & " and tipo = 'F'"
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) fornecedor(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Fornecedor(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With txtIDcliente
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    ProcLimpar
    ProcAtualizalista (1)
    Frame1.Enabled = False
    ProcLimparTudo
    Novo_Fornecedor = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_cobranca()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With listacobranca
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) local(is) para cobrança?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from clientes_cobranca where Idcobranca = " & .ListItems(InitFor)
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir local de cobrança"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & txtnomerazao
            Documento1 = "Endereço: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(is) de cobrança antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Local(is) de cobrança excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregalistacobranca
    ProcLimpacamposCobranca
    Frame16.Enabled = False
    Novo_Fornecedor3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_comercial()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_banco
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) banco(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Compras_fornecedores_banco where id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and id = " & .ListItems(InitFor) & " and tipo = 'F'"
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir banco"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & txtnomerazao
            Documento1 = "Banco: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) bancos(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Banco(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpacampos_banco
    Frame10.Enabled = False
    ProcCarregalista_banco
    Novo_Fornecedor5 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_entrega()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaEntrega
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) local(is) para entrega?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from clientes_entrega where Identrega = " & .ListItems(InitFor)
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir local de entrega"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & txtnomerazao
            Documento1 = "Endereço: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) local(is) para entrega antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Local(is) para entrega excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Proccarregalistaentrega
    ProcLimpacamposEntrega
    Frame11.Enabled = False
    Novo_Fornecedor2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_familia()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_familia
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) família(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from compras_fornecedores_familia where idfamilia = " & .ListItems(InitFor) & " and tipo = 'F'"
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir família"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & txtnomerazao
            Documento1 = "Família: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) família(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Família(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    cmbfamilia.ListIndex = -1
    txtid_familia = 0
    Frame6.Enabled = False
    ProcCarregaListaFamilia
    Novo_Fornecedor4 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_Segmento()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_Segmento
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) segmento(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            Conexao.Execute "DELETE from Compras_fornecedores_segmentos where id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir segmento"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & txtnomerazao
            Documento1 = "Segmento: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) segmento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Segmento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Segmento
    Frame5.Enabled = False
    ProcCarregaListaSegmento
    Novo_Fornecedor6 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_contato()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_contato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) contato(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from contatos_fornecedor where Idcontato = " & .ListItems(InitFor)
            '==================================
            Modulo = "Compras/Fornecedores"
            Evento = "Excluir contato"
            ID_documento = .ListItems(InitFor)
            Documento = "Fornecedor: " & txtnomerazao
            Documento1 = "Contato: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) contato(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Contato(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    procLimpacamposContatos
    ProcCarregaListaContatos
    Frame3.Enabled = False
    Novo_Fornecedor1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procLimpacamposContatos()
On Error GoTo tratar_erro

txtIDContato.Text = 0
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel1 = pubUsuario
txtNomeContato.Text = ""
txtdepartamento.Text = ""
txttelcontato.Text = ""
TxtEmail_Contato.Text = ""
Chk_enviar_NFe.Value = 0
Chk_enviar_boleto.Value = 0
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaContatos()
On Error GoTo tratar_erro

Lista_contato.ListItems.Clear
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Contatos_fornecedor where Idfornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " order by nome", Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    TBClientes.MoveLast
    ''PBLista.Min = 0
    ''PBLista.Max = TBClientes.RecordCount
    ''PBLista.Value = 1
    Contador = 0
    TBClientes.MoveFirst
    Do While TBClientes.EOF = False
        With Lista_contato.ListItems
            .Add = TBClientes!idcontato
            .Item(.Count).SubItems(1) = IIf(IsNull(TBClientes!Nome), "", (TBClientes!Nome))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBClientes!Departamento), "", (TBClientes!Departamento))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBClientes!ramal), "", (TBClientes!ramal))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBClientes!Email), "", (TBClientes!Email))
        End With
        TBClientes.MoveNext
        Contador = Contador + 1
        ''PBLista.Value = Contador
    Loop
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbUF_Click()
On Error GoTo tratar_erro

If cmbOrigem.Text = "Nacional" And cmbuf.Text = "EX" Or cmbOrigem = "Estrangeiro" And cmbuf <> "EX" And cmbuf <> "" Then
    cmbuf.ListIndex = -1
    cmbCidade.Clear
    txtCidade.Text = ""
ElseIf cmbuf.Text <> "EX" Then
        ProcCarregaComboCidade cmbCidade, "Sigla_UF = '" & cmbuf & "'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbuf_cobranca_Click()
On Error GoTo tratar_erro

If cmbOrigem.Text = "Nacional" And cmbuf_cobranca.Text = "EX" Or cmbOrigem = "Estrangeiro" And cmbuf_cobranca <> "EX" And cmbuf_cobranca <> "" Then
    cmbuf_cobranca.ListIndex = -1
    cmbCidade_cobranca.Clear
    txtcidade_cobranca.Text = ""
ElseIf cmbuf_cobranca <> "EX" Then
        ProcCarregaComboCidade cmbCidade_cobranca, "Sigla_UF = '" & cmbuf_cobranca & "'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbuf_entrega_Click()
On Error GoTo tratar_erro

If cmbOrigem.Text = "Nacional" And cmbuf_entrega.Text = "EX" Or cmbOrigem = "Estrangeiro" And cmbuf_entrega <> "EX" And cmbuf_entrega <> "" Then
    cmbuf_entrega.ListIndex = -1
    cmbCidade_Entrega.Clear
    txtCidade_Entrega.Text = ""
ElseIf cmbuf_entrega <> "EX" Then
        ProcCarregaComboCidade cmbCidade_Entrega, "Sigla_UF = '" & cmbuf_entrega & "'", False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_buscarCEP_Click()
On Error GoTo tratar_erro

If txtCEP = "" Then Exit Sub
If cmbOrigem = "" Or cmbOrigem = "Estrangeiro" Then
    USMsgBox ("Só é permitido carregar os dados pelo CEP se a origem for Nacional."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunBuscaEndereco(txtCEP) = False Then
    USMsgBox ("Não foi encontrado nenhuma informação pelo CEP informado."), vbExclamation, "CAPRIND v5.0"
    cmbTipo_endereco.ListIndex = -1
    txtendereco = ""
    cmbTipo_bairro.ListIndex = -1
    txtBairro = ""
    cmbuf.ListIndex = -1
    cmbCidade.ListIndex = -1
    Exit Sub
Else
    If USMsgBox("Deseja carregar os dados em maiúsculo?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido = True Else Permitido = False
    'cmbTipo_endereco = Tipo_endereco
    txtendereco = Trim(IIf(Permitido = True, UCase(Endereco), Endereco))
    txtBairro = Trim(IIf(Permitido = True, UCase(Bairro), Bairro))
    cmbuf = UF
    cmbCidade = Trim(FunTiraAcentosTexto(Cidade))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

Txt_caminho_certificado = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_CFOP_Click()
On Error GoTo tratar_erro

txtID_cfop = ""
txtCFOP = ""
txtOperacao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_tipo_dcto_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Clientes = False
Compras_Fornecedores = True
frmContas_Tipo_Dcto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If Txt_caminho_certificado <> "" Then ProcAbrirArquivo Txt_caminho_certificado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCertificado_Click()
On Error GoTo tratar_erro

Sit_REG = 2
frmSegmento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcfop_Click()
On Error GoTo tratar_erro

Compras_Fornecedores = True
Clientes = False
Vendas_Proposta = False
Vendas_PI = False
Faturamento = False
Compras_Pedido = False
Sit_REG = 2
frm_ListaNatureza.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCond_pag_padrao_Click()
On Error GoTo tratar_erro

Tipo = "C"
Aplic = 1
Clientes = False
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Compras_Fornecedores = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdConsultar_Click()
On Error GoTo tratar_erro
Dim resposta As String
Dim p As Object

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
'Debug.print CnpjNF
End If
TBAbrir.Close
If CnpjNF = "34270461000104" Then
CnpjNF = ReturnNumbersOnly("40.279.755/0001-91")
End If

resposta = consultarCadastroContribuinte(CnpjNF, cmbuf.Text, ReturnNumbersOnly(txtcnpj.Text), "CNPJ")
'Debug.print resposta
status = LerDadosJSON(resposta, "status", "", "")
   If status = "200" Then
      Set p = JSON.parse(resposta)
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "CNPJ da consulta nao cadastrado como contribuinte na UF. CNPJ: 16740838000151" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não é contribuinte do ICMS"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeição: CNPJ da consulta não cadastrado como contribuinte na UF" Then
      USMsgBox "Não será possivel buscar o cadastro desse cnpj pois não é contribuinte do ICMS"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: CNPJ da consulta nao cadastrado como contribuinte na UF" Then
      USMsgBox "Rejeicao: CNPJ da consulta nao cadastrado como contribuinte na UF"
      Exit Sub
      End If
      
      If p.Item("retConsCad").Item("infCons").Item("xMotivo") = "Rejeicao: Sigla da UF da consulta difere da UF do Web Service" Then
      USMsgBox "Rejeicao: Sigla da UF da consulta difere da UF do Web Service"
      Exit Sub
      End If
      
      
      '"Rejeicao: Sigla da UF da consulta difere da UF do Web Service"

      cmbuf.Text = Trim(p.Item("retConsCad").Item("infCons").Item("UF"))
      txtRG_IE = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("IE"))
      txtnomerazao.Text = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      'txtnomefantasia = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome"))
      txtendereco = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xLgr"))
      txtNumero = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("nro"))
      txtBairro = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xBairro"))
      txtCEP = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("CEP"))
      'cmbCidade.Text = Trim(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("ender").Item("xMun"))
      cmbRegimeTributario.Text = IIf(p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xRegApur") = "NORMAL - REGIME PERIÓDICO DE APURAÇÃO", "Lucro presumido", "Simples Nacional")
      'txtnomefantasia = p.Item("retConsCad").Item("infCons").Item("infCad").Item(1).Item("xNome")
      If txtCEP <> "" Then
      Cmd_buscarCEP_Click
      End If
      txtCategoria.Text = "A"
      USMsgBox "Consulta relizada com sucesso, dados carregados", vbInformation, "CAPRIND v5.0"
      
   Else
      USMsgBox resposta, vbCritical, "CAPRIND v5.0"
   End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho_certificado = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImpostos_padrao_Click()
On Error GoTo tratar_erro

Aplic = 7
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Compras_Fornecedores = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocTransp_Click()
On Error GoTo tratar_erro

With Cmb_tipo_transp
    If .Text = "Cliente" Then
        ProcConfVariaveisLocCliente False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
        frmVendas_LocalizarCliente.Show 1
    ElseIf .Text = "Fornecedor" Then
            ProcConfVariaveisLocForn False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
            FrmCompras_localizafornecedor.Show 1
        Else
            frmFaturamento_Prod_Serv_Localizar_Empresa.Show 1
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_cobranca()
On Error GoTo tratar_erro

If USMsgBox("Deseja aproveitar os dados principais do fornecedor?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    txtID_cobranca = 0
    If cmbPessoa = "JURÍDICA" Then txtCNPJ_cobranca = txtcnpj Else txtCNPJ_cobranca = txtCpf
    If cmbTipo_endereco <> "" Then cmbTipo_endereco_cobranca = cmbTipo_endereco
    txtendereco_cobranca = txtendereco
    txtNumero_cobranca = txtNumero
    txtComplemento_cobranca = txtComplemento
    If cmbTipo_bairro <> "" Then cmbTipo_bairro_cobranca = cmbTipo_bairro
    txtbairro_cobranca = txtBairro
    If cmbuf <> "EX" Then
        cmbuf_cobranca = cmbuf
        cmbCidade_cobranca.Visible = True
        txtcidade_cobranca.Visible = False
        cmbCidade_cobranca = cmbCidade
    Else
        cmbuf_cobranca = cmbuf
        cmbCidade_cobranca.Visible = False
        txtcidade_cobranca.Visible = True
        txtcidade_cobranca = txtCidade
    End If
    txtcxpostal_cobranca = txtcaixapostal
    mskcep_cobranca = txtCEP
    txttel1_cobranca = txtTelefones
    txtfax_cobranca = txtFax
    txtemail_cobranca.Text = txtEmail.Text
    txtSite_cobranca.Text = txtSite.Text
    CodigoLista3 = 0
Else
    ProcLimpacamposCobranca
End If
Frame16.Enabled = True
txtendereco_cobranca.SetFocus
Novo_Fornecedor3 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_comerciais()
On Error GoTo tratar_erro

ProcLimpacampos_banco
Novo_Fornecedor5 = True
Frame10.Enabled = True
txtBanco.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_Contato()
On Error GoTo tratar_erro

procLimpacamposContatos
Novo_Fornecedor1 = True
Frame3.Enabled = True
txtNomeContato.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpacamposEntrega()
On Error GoTo tratar_erro

txtID_entrega.Text = 0
txtData2 = Format(Date, "dd/mm/yy")
txtResponsavel2 = pubUsuario
If cmbPessoa <> "" Then
    If cmbPessoa = "JURÍDICA" Then txtCNPJ_entrega.Text = "__.___.___/____-__" Else txtCNPJ_entrega.Text = "___.___.___-__"
End If
cmbTipo_endereco_entrega.ListIndex = -1
txtEndereco_entrega.Text = ""
txtNumero_entrega.Text = ""
txtComplemento_entrega = ""
cmbTipo_bairro_entrega.ListIndex = -1
txtBairro_entrega.Text = ""
txtCidade_Entrega.Text = ""
cmbuf_entrega.ListIndex = -1
txtcxpostal_entrega.Text = ""
mskcep_entrega.Text = ""
txttel1_entrega.Text = ""
txttel2_entrega.Text = ""
txttel3_entrega.Text = ""
txttel4_entrega.Text = ""
txtfax_entrega.Text = ""
txtemail_entrega.Text = ""
txtSite_cobranca.Text = ""
CodigoLista2 = 0
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpacamposCobranca()
On Error GoTo tratar_erro

txtID_cobranca.Text = 0
txtData3 = Format(Date, "dd/mm/yy")
txtResponsavel3 = pubUsuario
If cmbPessoa <> "" Then
    If cmbPessoa = "JURÍDICA" Then txtCNPJ_cobranca = "__.___.___/____-__" Else txtCNPJ_cobranca = "___.___.___-__"
End If
cmbTipo_endereco_cobranca.ListIndex = -1
txtendereco_cobranca.Text = ""
txtNumero_cobranca.Text = ""
txtComplemento_cobranca = ""
cmbTipo_bairro_cobranca.ListIndex = -1
txtbairro_cobranca.Text = ""
txtcidade_cobranca.Text = ""
cmbuf_cobranca.ListIndex = -1
txtcxpostal_cobranca.Text = ""
mskcep_cobranca = ""
txttel1_cobranca.Text = ""
txttel2_cobranca.Text = ""
txttel3_cobranca.Text = ""
txttel4_cobranca.Text = ""
txtfax_cobranca.Text = ""
txtemail_cobranca.Text = ""
txtSite_cobranca.Text = ""
CodigoLista3 = 0
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_entrega()
On Error GoTo tratar_erro

If USMsgBox("Deseja aproveitar os dados principais do fornecedor?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    txtID_entrega = 0
    If cmbPessoa = "JURÍDICA" Then txtCNPJ_entrega = txtcnpj Else txtCNPJ_entrega = txtCpf
    If cmbTipo_endereco <> "" Then cmbTipo_endereco_entrega = cmbTipo_endereco
    txtEndereco_entrega = txtendereco
    txtNumero_entrega = txtNumero
    txtComplemento_entrega = txtComplemento
    If cmbTipo_bairro <> "" Then cmbTipo_bairro_entrega = cmbTipo_bairro
    txtBairro_entrega = txtBairro
    If cmbuf <> "EX" Then
        cmbuf_entrega = cmbuf
        cmbCidade_Entrega.Visible = True
        txtCidade_Entrega.Visible = False
        cmbCidade_Entrega = cmbCidade
    Else
        cmbuf_entrega = cmbuf
        cmbCidade_Entrega.Visible = False
        txtCidade_Entrega.Visible = True
        txtCidade_Entrega = txtCidade
    End If
    txtcxpostal_entrega = txtcaixapostal
    mskcep_entrega = txtCEP
    txttel1_entrega = txtTelefones
    txtfax_entrega = txtFax
    txtemail_entrega.Text = txtEmail.Text
    txtsite_entrega.Text = txtSite.Text
    CodigoLista3 = 0
Else
    ProcLimpacamposEntrega
End If
Frame11.Enabled = True
txtEndereco_entrega.SetFocus
Novo_Fornecedor2 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_familia()
On Error GoTo tratar_erro
  
cmbfamilia.ListIndex = -1
txtid_familia = 0
Novo_Fornecedor4 = True
Frame6.Enabled = True
cmbfamilia.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_segmento()
On Error GoTo tratar_erro
  
ProcLimpaCampos_Segmento
Novo_Fornecedor6 = True
Frame5.Enabled = True
cmdSegmento_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDcliente = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores order by Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDCliente = " & txtIDcliente)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtIDcliente = TBLISTA!IDCliente
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from Compras_fornecedores where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        procLimpacamposContatos
        ProcLimpacamposEntrega
        ProcLimpacamposCobranca
        ProcLimpaFamilia
        ProcLimpacampos_banco
        ProcLimpaCampos_Outros
        ProcLimpaCampos_Segmento
        ProcPuxaDados
        ProcCarregaListaContatos
        Proccarregalistaentrega
        ProcCarregalistacobranca
        ProcCarregaListaFamilia
        ProcCarregalista_banco
        ProcCarregaListaSegmento
        ProcPuxaDados_Outros
    Else
        USMsgBox ("Fim dos cadastros de fornecedor."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Fornecedor1 = False
Novo_Fornecedor4 = False
Novo_Fornecedor5 = False
Novo_Fornecedor6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_cobranca()
On Error GoTo tratar_erro

If Frame16.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbPessoa = "JURÍDICA" And txtCNPJ_cobranca.Text <> "__.___.___/____-__" Then
    If Funconsistir_CgcCpf(txtCNPJ_cobranca) = False Then
        USMsgBox ("O número do CNPJ digitado não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
        txtCNPJ_cobranca.SetFocus
        Exit Sub
    End If
End If
Acao = "salvar"
If txtendereco_cobranca <> "" And cmbOrigem = "Nacional" And cmbTipo_endereco_cobranca = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_endereco_cobranca.SetFocus
    Exit Sub
End If
If txtendereco_cobranca.Text = "" Then
    NomeCampo = "o endereço"
    ProcVerificaAcao
    txtendereco_cobranca.SetFocus
    Exit Sub
End If
If txtbairro_cobranca <> "" And cmbOrigem = "Nacional" And cmbTipo_bairro_cobranca = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_bairro_cobranca.SetFocus
    Exit Sub
End If
If txtcidade_cobranca <> "" And cmbuf_cobranca <> "" And cmbuf_cobranca <> "EX" Then
    If cmbOrigem = "Nacional" And FunVerificaCidade(txtcidade_cobranca, cmbuf_cobranca) = False Then Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from clientes_cobranca where idcobranca = " & txtID_cobranca, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "este endereço de cobrança", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDCliente = txtIDcliente
If txtData3 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData3
If txtResponsavel3 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel3
TBGravar!Tipo = "F"
TBGravar!CNPJ = txtCNPJ_cobranca.Text
TBGravar!Tipo_endereco = cmbTipo_endereco_cobranca
TBGravar!endereco_Cobranca = txtendereco_cobranca
TBGravar!Numero = txtNumero_cobranca.Text
TBGravar!complemento = txtComplemento_cobranca
TBGravar!Tipo_bairro = cmbTipo_bairro_cobranca
TBGravar!bairro_Cobranca = txtbairro_cobranca
If cmbuf_cobranca <> "EX" Then TBGravar!cidade_Cobranca = IIf(cmbCidade_cobranca.Text = "", Null, cmbCidade_cobranca.Text) Else TBGravar!cidade_Cobranca = IIf(txtcidade_cobranca.Text = "", Null, txtcidade_cobranca.Text)
TBGravar!uf_Cobranca = cmbuf_cobranca
TBGravar!cxpostal_Cobranca = txtcxpostal_cobranca
TBGravar!cep_Cobranca = mskcep_cobranca
TBGravar!tel1_Cobranca = txttel1_cobranca
TBGravar!tel2_Cobranca = txttel2_cobranca
TBGravar!tel3_Cobranca = txttel3_cobranca
TBGravar!tel4_Cobranca = txttel4_cobranca
TBGravar!fax_Cobranca = txtfax_cobranca
TBGravar!email_Cobranca = IIf(txtemail_cobranca.Text = "", Null, LCase(txtemail_cobranca.Text))
TBGravar!site_cobranca = IIf(txtSite_cobranca.Text = "", Null, LCase(txtSite_cobranca.Text))
TBGravar.Update
txtID_cobranca = TBGravar!idCobranca
TBGravar.Close
ProcCarregalistacobranca
If Novo_Fornecedor3 = True Then
    USMsgBox ("Novo local para cobrança cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo local para cobrança"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar local para cobrança"
    If listacobranca.ListItems.Count <> 0 And CodigoLista3 <> 0 Then
        listacobranca.SelectedItem = listacobranca.ListItems(CodigoLista3)
        listacobranca.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Fornecedores"
ID_documento = txtID_cobranca
Documento = "Fornecedor: " & txtnomerazao
Documento1 = "Endereço: " & txtendereco_cobranca
ProcGravaEvento
'==================================
Novo_Fornecedor3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_comerciais()
On Error GoTo tratar_erro

If Frame10.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtBanco = "" Then
    NomeCampo = "o banco"
    ProcVerificaAcao
    txtBanco.SetFocus
    Exit Sub
End If
If txtAgencia = "" Then
    NomeCampo = "a agência"
    ProcVerificaAcao
    txtAgencia.SetFocus
    Exit Sub
End If
If txtConta = "" Then
    NomeCampo = "a conta"
    ProcVerificaAcao
    txtConta.SetFocus
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_fornecedores_banco where banco = '" & txtBanco & "' and id <> " & txtID_banco & " and id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'F'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Este banco já está cadastrado."), vbExclamation, "CAPRIND v5.0"
    txtBanco.SetFocus
    Exit Sub
End If
TBAbrir.Close
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_fornecedores_banco where id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and id = " & txtID_banco & " and tipo = 'F'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "estes dados bancários", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!id_fornecedor = txtIDcliente
If txtData5 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData5
If txtResponsavel5 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel5
TBGravar!Tipo = "F"
TBGravar!Banco = txtBanco
TBGravar!Agencia = txtAgencia
TBGravar!Conta = txtConta
TBGravar.Update
txtID_banco = TBGravar!ID
TBGravar.Close
ProcCarregalista_banco
If Novo_Fornecedor5 = True Then
    USMsgBox ("Novo banco cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo banco"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar banco"
    If lista_banco.ListItems.Count <> 0 And CodigoLista5 <> 0 Then
        lista_banco.SelectedItem = lista_banco.ListItems(CodigoLista5)
        lista_banco.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Fornecedores"
ID_documento = txtID_banco
Documento = "Fornecedor: " & txtnomerazao
Documento1 = "Banco: " & txtBanco
ProcGravaEvento
'==================================
Novo_Fornecedor5 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalistaentrega()
On Error GoTo tratar_erro

ListaEntrega.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_entrega where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'F' order by endereco_entrega", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEntrega.ListItems
            .Add = TBLISTA!identrega
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!endereco_entrega), "", TBLISTA!endereco_entrega)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!bairro_entrega), "", TBLISTA!bairro_entrega)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!cidade_entrega), "", TBLISTA!cidade_entrega)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!uf_entrega), "", TBLISTA!uf_entrega)
            TBLISTA.MoveNext
            Contador = Contador + 1
            'PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregalistacobranca()
On Error GoTo tratar_erro

listacobranca.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_cobranca where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'F' order by endereco_cobranca", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With listacobranca.ListItems
            .Add = TBLISTA!idCobranca
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!endereco_Cobranca), "", TBLISTA!endereco_Cobranca)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!bairro_Cobranca), "", TBLISTA!bairro_Cobranca)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!uf_Cobranca), "", TBLISTA!uf_Cobranca)
            TBLISTA.MoveNext
            Contador = Contador + 1
            'PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_contato()
On Error GoTo tratar_erro

If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtNomeContato.Text = "" Then
    USMsgBox ("Informe o nome do contato antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtNomeContato.SetFocus
    Exit Sub
End If
If txttelcontato = "" And TxtEmail_Contato.Text = "" Then
    USMsgBox ("Informe o telefone ou email do contato antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Contatos_fornecedor where idcontato = " & txtIDContato.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    TBAbrir.AddNew
Else
    If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "este contato", "alterar", True, True) = False Then Exit Sub
End If
ProcEnviaDadosContato
TBAbrir.Update
txtIDContato = TBAbrir!idcontato
TBAbrir.Close
ProcCarregaListaContatos
If Novo_Fornecedor1 = True Then
    USMsgBox ("Novo contato cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo contato"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar contato"
    If CodigoLista1 <> 0 And Lista_contato.ListItems.Count <> 0 Then
        Lista_contato.SelectedItem = Lista_contato.ListItems(CodigoLista1)
        Lista_contato.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Fornecedores"
ID_documento = txtIDContato
Documento = "Fornecedor: " & txtnomerazao
Documento1 = "Contato: " & txtNomeContato
ProcGravaEvento
'==================================
Novo_Fornecedor1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosContato()
On Error GoTo tratar_erro

TBAbrir!IDFornecedor = txtIDcliente
If txtData1 = "" Then TBAbrir!Data = Date Else TBAbrir!Data = txtData1
If txtResponsavel1 = "" Then TBAbrir!Responsavel = pubUsuario Else TBAbrir!Responsavel = txtResponsavel1
TBAbrir!Nome = txtNomeContato
TBAbrir!Departamento = txtdepartamento
TBAbrir!ramal = txttelcontato
TBAbrir!Email = IIf(TxtEmail_Contato.Text = "", Null, LCase(TxtEmail_Contato.Text))
If Chk_enviar_NFe.Value = 1 Then TBAbrir!Enviar_NFe = True Else TBAbrir!Enviar_NFe = False
If Chk_enviar_boleto.Value = 1 Then TBAbrir!Enviar_boleto = True Else TBAbrir!Enviar_boleto = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_entrega()
On Error GoTo tratar_erro

If Frame11.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbPessoa = "JURÍDICA" And txtCNPJ_entrega.Text <> "__.___.___/____-__" Then
    If Funconsistir_CgcCpf(txtCNPJ_entrega) = False Then
        USMsgBox ("O número do CNPJ digitado não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
        txtCNPJ_entrega.SetFocus
        Exit Sub
    End If
End If
Acao = "salvar"
If txtEndereco_entrega <> "" And cmbOrigem = "Nacional" And cmbTipo_endereco_entrega = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_endereco_entrega.SetFocus
    Exit Sub
End If
If txtEndereco_entrega.Text = "" Then
    NomeCampo = "o endereço"
    ProcVerificaAcao
    txtEndereco_entrega.SetFocus
    Exit Sub
End If
If txtBairro_entrega <> "" And cmbOrigem = "Nacional" And cmbTipo_bairro_entrega = "" Then
    NomeCampo = "o tipo"
    ProcVerificaAcao
    cmbTipo_bairro_entrega.SetFocus
    Exit Sub
End If
If txtCidade_Entrega <> "" And cmbuf_entrega <> "" And cmbuf_entrega <> "EX" Then
    If cmbOrigem = "Nacional" And FunVerificaCidade(txtCidade_Entrega, cmbuf_entrega) = False Then Exit Sub
End If
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes_entrega where identrega = " & txtID_entrega, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew
Else
    If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "este endereço de entrega", "alterar", True, True) = False Then Exit Sub
End If
If txtData2 = "" Then TBClientes!Data = Date Else TBClientes!Data = txtData2
If txtResponsavel2 = "" Then TBClientes!Responsavel = pubUsuario Else TBClientes!Responsavel = txtResponsavel2
TBClientes!IDCliente = txtIDcliente
TBClientes!Tipo = "F"
TBClientes!CNPJ = txtCNPJ_entrega.Text
TBClientes!Tipo_endereco = cmbTipo_endereco_entrega
TBClientes!endereco_entrega = txtEndereco_entrega
TBClientes!Numero = txtNumero_entrega.Text
TBClientes!complemento = txtComplemento_entrega
TBClientes!Tipo_bairro = cmbTipo_bairro_entrega
TBClientes!bairro_entrega = txtBairro_entrega
If cmbuf_entrega <> "EX" Then TBClientes!cidade_entrega = IIf(cmbCidade_Entrega.Text = "", Null, cmbCidade_Entrega.Text) Else TBClientes!cidade_entrega = IIf(txtCidade_Entrega.Text = "", Null, txtCidade_Entrega.Text)
TBClientes!uf_entrega = cmbuf_entrega
TBClientes!cxpostal_entrega = txtcxpostal_entrega
TBClientes!cep_entrega = mskcep_entrega
TBClientes!tel1_entrega = txttel1_entrega
TBClientes!tel2_entrega = txttel2_entrega
TBClientes!tel3_entrega = txttel3_entrega
TBClientes!tel4_entrega = txttel4_entrega
TBClientes!fax_entrega = txtfax_entrega
TBClientes!email_entrega = IIf(txtemail_entrega.Text = "", Null, LCase(txtemail_entrega.Text))
TBClientes!Site_Entrega = IIf(txtsite_entrega.Text = "", Null, LCase(txtsite_entrega.Text))
TBClientes.Update
txtID_entrega = TBClientes!identrega
TBClientes.Close
Proccarregalistaentrega
If Novo_Fornecedor2 = True Then
    USMsgBox ("Novo local para entrega cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo local para entrega"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar local para entrega"
    If ListaEntrega.ListItems.Count <> 0 And CodigoLista2 <> 0 Then
        ListaEntrega.SelectedItem = ListaEntrega.ListItems(CodigoLista2)
        ListaEntrega.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Fornecedores"
ID_documento = txtID_entrega
Documento = "Fornecedor: " & txtnomerazao
Documento1 = "Endereço: " & txtEndereco_entrega
ProcGravaEvento
'==================================
Novo_Fornecedor2 = False
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_familia()
On Error GoTo tratar_erro

If Frame6.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbfamilia = "" Then
    USMsgBox ("Informe a família antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmbfamilia.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from compras_fornecedores_familia where idfamilia = " & txtid_familia, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "esta família", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDCliente = txtIDcliente
If txtData4 = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData4
If txtResponsavel4 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel4
TBGravar!Tipo = "F"
TBGravar!Familia = cmbfamilia
TBGravar.Update
txtid_familia = TBGravar!idFamilia
TBGravar.Close
ProcCarregaListaFamilia
If Novo_Fornecedor4 = True Then
    USMsgBox ("Nova família cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova família"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar família"
    If lista_familia.ListItems.Count <> 0 And CodigoLista4 <> 0 Then
        lista_familia.SelectedItem = lista_familia.ListItems(CodigoLista4)
        lista_familia.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Fornecedores"
ID_documento = txtid_familia
Documento = "Fornecedor: " & txtnomerazao
Documento1 = "Família: " & cmbfamilia
ProcGravaEvento
'==================================
Novo_Fornecedor4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_segmento()
On Error GoTo tratar_erro

If Frame5.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtSegmento = "" Then
    USMsgBox ("Informe o segmento antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmdSegmento_Click
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_fornecedores_segmentos where id = " & txtID_segmento, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "este segmento", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDFornecedor = txtIDcliente
If txtData_segmento = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData_segmento
If txtResp_segmento = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResp_segmento
TBGravar!Segmento = txtSegmento
TBGravar.Update
txtID_segmento = TBGravar!ID
TBGravar.Close
ProcCarregaListaSegmento
If Novo_Fornecedor6 = True Then
    USMsgBox ("Novo segmento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo segmento"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar segmento"
    If Lista_Segmento.ListItems.Count <> 0 And CodigoLista6 <> 0 Then
        Lista_Segmento.SelectedItem = Lista_Segmento.ListItems(CodigoLista6)
        Lista_Segmento.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Fornecedores"
ID_documento = txtID_segmento
Documento = "Fornecedor: " & txtnomerazao
Documento1 = "Família: " & txtSegmento
ProcGravaEvento
'==================================
Novo_Fornecedor6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_outros()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "o mesmo", "alterar", True, True) = False Then Exit Sub

Acao = "salvar"
If Chk_certificado.Value = 1 Or Chk_avaliado.Value = 1 Or chkAprovada_Cliente.Value = 1 Then
    If Chk_certificado.Value = 1 Then
        Texto = "certificação"
    ElseIf Chk_avaliado.Value = 1 Then
            Texto = "avaliação"
        Else
            Texto = "fonte aprovada pelo cliente"
    End If
    If IsDate(txtdata_venc) = False Then
        NomeCampo = "a data de vencimento da " & Texto
        ProcVerificaAcao
        txtdata_venc.SetFocus
        Exit Sub
    End If
End If
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from compras_fornecedores where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    TBFornecedor!Obs = txtObs
    If Chk_certificado.Value = 1 Then
        TBFornecedor!Fornecedor = "C"
    ElseIf Chk_avaliado.Value = 1 Then
            TBFornecedor!Fornecedor = "A"
        ElseIf chkAprovada_Cliente.Value = 1 Then
                TBFornecedor!Fornecedor = "F"
            Else
                TBFornecedor!Fornecedor = Null
    End If
    If Frame7.Enabled = True Then TBFornecedor!Data_venc = txtdata_venc Else TBFornecedor!Data_venc = Null
    TBFornecedor!ICMS_ind = IIf(Txt_ICMS_ind = "", 0, Txt_ICMS_ind)
    If chkSedex.Value = 1 Then TBFornecedor!Sedex = True Else TBFornecedor!Sedex = False
    If chkICMSST.Value = 1 Then TBFornecedor!SimplesICMSST = True Else TBFornecedor!SimplesICMSST = False
    If chkDesignado.Value = 1 Then TBFornecedor!Designado = True Else TBFornecedor!Designado = False
    TBFornecedor!Banco = IIf(cmbBanco = "", Null, cmbBanco)
    TBFornecedor!Tipo_doc = IIf(cmbTipo_doc = "", Null, cmbTipo_doc)
    TBFornecedor!Certificado = IIf(txtCertificado = "", Null, txtCertificado)
    TBFornecedor!Caminho_certificado = Txt_caminho_certificado
    TBFornecedor.Update
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End If
TBFornecedor.Close

'==================================
Modulo = "Compras/Fornecedores"
Evento = "Alterar"
ID_documento = txtIDcliente
Documento = "Fornecedor: " & txtnomerazao
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Fornecedor.AbsolutePage <> 2 Then
    If TBLISTA_Fornecedor.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Fornecedor.PageCount - 1)
    Else
        TBLISTA_Fornecedor.AbsolutePage = TBLISTA_Fornecedor.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Fornecedor.AbsolutePage)
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
    TBLISTA_Fornecedor.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Fornecedor.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Fornecedor.AbsolutePage = 1
ProcExibePagina (TBLISTA_Fornecedor.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Fornecedor.AbsolutePage <> -3 Then
    If TBLISTA_Fornecedor.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Fornecedor.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Fornecedor.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Fornecedor.AbsolutePage = TBLISTA_Fornecedor.PageCount
ProcExibePagina (TBLISTA_Fornecedor.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSegmento_Click()
On Error GoTo tratar_erro

Sit_REG = 1
frmSegmento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdTransporte_padrao_Click()
On Error GoTo tratar_erro
Tipo = "C"
Aplic = 6
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Compras_Fornecedores = True

Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdValidade_Padrao_Click()
On Error GoTo tratar_erro

Aplic = 10
Compras_Cotacao = False
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Compras_Fornecedores = True
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

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
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: procFiltrar_todos
            Case vbKeyF8: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF10:  If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "Compras/Fornecedores"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case Is <> 0:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoTab
            Case vbKeyF3: ProcSalvarTab
            Case vbKeyF4: ProcExcluirTab
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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15200, 15, True
ProcCarregaToolBar2 Me, 15200, 10, True
ProcCarregaToolBar3 Me, 15200, 8, True
Formulario = "Compras/Fornecedores"
Direitos
SSTab1.Tab = 0
ProcCarregaCombos
ProcLimpaVariaveisPrincipais
USToolBar2.Visible = False
Cmb_opcao_lista.Text = "Validação"

ProcRemoveObjetosResize Me

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
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbOrigem = "Nacional" And cmbPessoa = "JURÍDICA" Then
    If cmbRegimeTributario = "" Then
        If USMsgBox("O fornecedor será cadastrado sem regime tributário, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
    End If
End If

Acao = "salvar"
If txtnomerazao = "" Then
    NomeCampo = "a razão"
    ProcVerificaAcao
    txtnomerazao.SetFocus
    Exit Sub
End If
If Chk_prospecto.Value = 0 And Chk_enviar_NF.Value = 1 Then
    If cmbPessoa.Text = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        If Frame1.Enabled = True Then cmbPessoa.SetFocus
        Exit Sub
    End If
    If txtCategoria.Text = "" Then
        NomeCampo = "o IQF"
        ProcVerificaAcao
        txtCategoria.SetFocus
        Exit Sub
    End If
    If cmbOrigem.Text = "" Then
        NomeCampo = "a origem"
        ProcVerificaAcao
        cmbOrigem.SetFocus
        Exit Sub
    End If
    If cmbOrigem = "Nacional" Then
        If Left(cmbPessoa, 8) = "JURÍDICA" And txtcnpj = "__.___.___/____-__" Then
            NomeCampo = "o CNPJ"
            ProcVerificaAcao
            txtcnpj.SetFocus
            Exit Sub
        ElseIf Left(cmbPessoa, 6) = "FÍSICA" And txtCpf.Text = "___.___.___-__" Then
                NomeCampo = "o CPF"
                ProcVerificaAcao
                txtCpf.SetFocus
                Exit Sub
        End If
    End If
    
    If Txt_pais = "" Then
        NomeCampo = "o país"
        ProcVerificaAcao
        Txt_pais.SetFocus
        Exit Sub
    End If
    If txtendereco <> "" And cmbOrigem = "Nacional" And cmbTipo_endereco = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        cmbTipo_endereco.SetFocus
        Exit Sub
    End If
    If txtendereco = "" Then
        NomeCampo = "o endereço"
        ProcVerificaAcao
        txtendereco.SetFocus
        Exit Sub
    End If
    If txtNumero.Text = "" Then
        NomeCampo = "o número"
        ProcVerificaAcao
        txtNumero.SetFocus
        Exit Sub
    End If
    If txtBairro <> "" And cmbOrigem = "Nacional" And cmbTipo_bairro = "" Then
        NomeCampo = "o tipo"
        ProcVerificaAcao
        cmbTipo_bairro.SetFocus
        Exit Sub
    End If
    If txtBairro = "" Then
        NomeCampo = "o bairro"
        ProcVerificaAcao
        txtBairro.SetFocus
        Exit Sub
    End If
    If cmbOrigem = "Nacional" Then
        If cmbuf.Text = "" Then
            NomeCampo = "o estado"
            ProcVerificaAcao
            cmbuf.SetFocus
            Exit Sub
        End If
        If cmbuf.Text = "EX" Then
            USMsgBox ("Não é permitido informar UF de exportação para fornecedor nacional."), vbInformation, "CAPRIND v5.0"
            cmbuf.SetFocus
            Exit Sub
        End If
        If cmbCidade = "" Then
            NomeCampo = "a cidade"
            ProcVerificaAcao
            cmbCidade.SetFocus
            Exit Sub
        End If
        If FunVerificaCidade(cmbCidade, cmbuf) = False Then Exit Sub
        
        If txtCEP.Text = "" Then
            NomeCampo = "o CEP"
            ProcVerificaAcao
            txtCEP.SetFocus
            Exit Sub
        End If
    End If
End If

If txtcnpj.Text <> "__.___.___/____-__" Then
    If Funconsistir_CgcCpf(txtcnpj) = False Then
        USMsgBox ("O número do CNPJ digitado não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
        txtcnpj.SetFocus
        Exit Sub
    End If
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select * from Compras_fornecedores where cpf_cnpj = '" & txtcnpj.Text & "' and idcliente <> " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        If Novo_Fornecedor = True Or TBFornecedor!Nome_Razao <> txtnomerazao.Text Then
            USMsgBox ("Já existe cadastro deste CNPJ para o fornecedor " & TBFornecedor!Nome_Razao & ", favor alterar o número do CNPJ."), vbExclamation, "CAPRIND v5.0"
            txtcnpj.Text = "__.___.___/____-__"
            txtcnpj.SetFocus
            TBFornecedor.Close
            Exit Sub
        End If
    End If
ElseIf txtCpf.Text <> "___.___.___-__" Then
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from Compras_fornecedores where cpf_cnpj = '" & txtCpf.Text & "' and idcliente <> " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then
            If Novo_Fornecedor = True Or TBFornecedor!Nome_Razao <> txtnomerazao.Text Then
                USMsgBox ("Já existe cadastro deste CPF para o fornecedor " & TBFornecedor!Nome_Razao & ", favor alterar o número do CPF."), vbExclamation, "CAPRIND v5.0"
                txtCpf = "___.___.___-__"
                txtCpf.SetFocus
                TBFornecedor.Close
                Exit Sub
            End If
        End If
End If

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from compras_fornecedores where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = True Then
    TBFornecedor.AddNew
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_fornecedores order by idcliente", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir.MoveLast
        IDCliente = TBAbrir!IDCliente + 1
    Else
        IDCliente = 1
    End If
    TBAbrir.Close
    txtIDcliente.Text = IDCliente
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "o fornecedor", True) = False Then Exit Sub
    If txtnomerazao <> TBFornecedor!Nome_Razao Or txtCategoria <> TBFornecedor!Categoria Or txtcnpj <> TBFornecedor!CPF_CNPJ Or txtRG_IE <> TBFornecedor!RG_IE Then
        If USMsgBox("Deseja atualizar os dados deste fornecedor em todos os módulos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            
            If txtnomerazao <> TBFornecedor!Nome_Razao Or txtCategoria <> TBFornecedor!Categoria Then Conexao.Execute "Update Compras_pedido Set idfornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & ", Fornecedor = '" & txtnomerazao.Text & "', Categoria = '" & txtCategoria.Text & "' where idfornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente)
            
            If txtnomerazao <> TBFornecedor!Nome_Razao Then
                Conexao.Execute "Update Cotacao_fornecedor Set forn = '" & txtnomerazao & "' where idforn = " & IIf(txtIDcliente = "", 0, txtIDcliente)
                Conexao.Execute "Update CQ_RNC Set Cliente_forn = '" & txtnomerazao & "' where ID_forn = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'F'"
                Conexao.Execute "Update item_aplicacoes Set aplicacao = '" & txtnomerazao & "' where ID_cliente_forn = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'F'"
                Conexao.Execute "Update tbl_ContasPagar Set txt_Fornecedor = '" & txtnomerazao.Text & "' where int_codforn = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'FO'"
                
                'Transportadora
                Conexao.Execute "Update Compras_fornecedores Set Transportadora = '" & txtnomerazao & "' where IDTransp = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo_transp = 'F'"
                Conexao.Execute "Update vendas_comercial Set Transportadora = '" & txtnomerazao & "' where IDInttransp = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo_transp = 'F'"
                Conexao.Execute "Update tbl_Dados_Transp Set txt_Razao = '" & txtnomerazao & "' where IdIntTransp = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo_transp = 'F'"
                
            End If
                        
            If txtnomerazao <> TBFornecedor!Nome_Razao Or txtcnpj <> TBFornecedor!CPF_CNPJ Or txtRG_IE <> TBFornecedor!RG_IE Then
                If TBFornecedor!Pessoa = "JURÍDICA" Then Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set txt_Razao_Nome = '" & txtnomerazao & "', txt_CNPJ_CPF = '" & txtcnpj & "', txt_IE_Cliente = '" & txtRG_IE & "' where Id_Int_Cliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and txt_Razao_Nome = '" & TBFornecedor!Nome_Razao & "' and Imprimir = 'False'"
                If TBFornecedor!Pessoa = "FÍSICA" Then Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set txt_Razao_Nome = '" & txtnomerazao & "', txt_CNPJ_CPF = '" & txtCpf & "' where Id_Int_Cliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and txt_Razao_Nome = '" & TBFornecedor!Nome_Razao & "' and Imprimir = 'False'"
            End If
        End If
    End If
End If
ProcEnviaDados
TBFornecedor.Update
TBFornecedor.Close

If Novo_Fornecedor = True Then
    USMsgBox "Novo fornecedor cadastrado com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Fornecedor = "Select * from Compras_fornecedores where idcliente = " & txtIDcliente
    ProcAtualizalista (1)
    
    Novo_Fornecedor = False
    USMsgBox "Não se esqueça de cadastrar o(s) contato(s) para este fornecedor.", vbInformation, "CAPRIND v5.0"
    SSTab1.Tab = 1
Else
    USMsgBox "Alteração efetuada com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Compras/Fornecedores"
    ID_documento = txtIDcliente
    Documento = "Fornecedor: " & txtnomerazao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    With txtIDcliente
        .Locked = False
        .TabStop = True
    End With
'    With cmbPessoa
'        .Locked = True
'        .TabStop = False
'    End With
    Novo_Fornecedor = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSql_Fornecedor = "" Then Exit Sub
Set TBLISTA_Fornecedor = CreateObject("adodb.recordset")
TBLISTA_Fornecedor.Open StrSql_Fornecedor, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Fornecedor.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Fornecedor.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Fornecedor.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Fornecedor.PageSize
ContadorReg = 1

'PBLista.Min = 0
'PBLista.Max = FunVerifMax'PBListaPaginacao(TBLISTA_Fornecedor.RecordCount - IIf(Pagina > 1, (TBLISTA_Fornecedor.PageSize * (Pagina - 1)), 0), TBLISTA_Fornecedor.PageSize)
'PBLista.Value = 1
Contador = 0
Do While TBLISTA_Fornecedor.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Fornecedor!IDCliente
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Fornecedor!DTCadastro), "", Format(TBLISTA_Fornecedor!DTCadastro, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Fornecedor!Responsavel), "", TBLISTA_Fornecedor!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Fornecedor!Nome_Razao), "", TBLISTA_Fornecedor!Nome_Razao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Fornecedor!DtValidacao), "Não", "Sim")
        .Item(.Count).SubItems(5) = TBLISTA_Fornecedor!ID
    End With
    TBLISTA_Fornecedor.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    'PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Fornecedor.RecordCount
If TBLISTA_Fornecedor.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Fornecedor.PageCount
ElseIf TBLISTA_Fornecedor.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Fornecedor.PageCount & " de: " & TBLISTA_Fornecedor.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Fornecedor.AbsolutePage - 1 & " de: " & TBLISTA_Fornecedor.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro
  
TBFornecedor!IDCliente = txtIDcliente.Text
If txtData.Text = "" Then TBFornecedor!DTCadastro = Date Else TBFornecedor!DTCadastro = txtData
If txtResponsavel = "" Then TBFornecedor!Responsavel = pubUsuario Else TBFornecedor!Responsavel = txtResponsavel
TBFornecedor!status = txtStatus
TBFornecedor!Nome_Razao = Replace(txtnomerazao, "'", " ")
TBFornecedor!NomeFantasia = IIf(txtnomefantasia = "", Null, txtnomefantasia)
If Cmb_centro <> "" Then TBFornecedor!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex) Else TBFornecedor!ID_CC = Null
If Chk_nao_contribuinte_ICMS.Value = 1 Then TBFornecedor!Nao_contribuinte_ICMS = True Else TBFornecedor!Nao_contribuinte_ICMS = False
TBFornecedor!Endereco = IIf(txtendereco = "", Null, txtendereco)
TBFornecedor!Bairro = IIf(txtBairro = "", Null, txtBairro)
If cmbuf = "EX" Then TBFornecedor!Cidade = IIf(txtCidade.Text = "", Null, txtCidade.Text) Else TBFornecedor!Cidade = IIf(cmbCidade.Text = "", Null, cmbCidade.Text)
If cmbRegimeTributario = "Lucro presumido" Then TBFornecedor!Presumido = True Else TBFornecedor!Presumido = False
If cmbRegimeTributario = "Simples nacional" Then TBFornecedor!Simples = True Else TBFornecedor!Simples = False
If cmbRegimeTributario = "Lucro real" Then TBFornecedor!Real = True Else TBFornecedor!Real = False
If cmbRegimeTributario = "MEI" Then TBFornecedor!MEI = True Else TBFornecedor!MEI = False

If cmbtransportadora <> "" Then
    Select Case Cmb_tipo_transp
        Case "Cliente": TBFornecedor!Tipo_transp = "C"
        Case "Fornecedor": TBFornecedor!Tipo_transp = "F"
        Case "Empresa": TBFornecedor!Tipo_transp = "E"
    End Select
    TBFornecedor!Transportadora = cmbtransportadora
    TBFornecedor!idTransp = cmbtransportadora.ItemData(cmbtransportadora.ListIndex)
Else
    TBFornecedor!Tipo_transp = ""
    TBFornecedor!Transportadora = ""
    TBFornecedor!idTransp = 0
End If

TBFornecedor!Telefones = IIf(txtTelefones.Text = "", Null, txtTelefones.Text)
TBFornecedor!RG_IE = IIf(txtRG_IE.Text = "", Null, txtRG_IE.Text)

If cmbOrigem = "Nacional" Then
    TBFornecedor!CPF_CNPJ = IIf(cmbPessoa = "JURÍDICA", txtcnpj.Text, txtCpf.Text)
ElseIf Novo_Fornecedor = True Then
        Contador = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_fornecedores where idTipoEmpresa = 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then Contador = TBAbrir.RecordCount
    
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from clientes where idTipoEmpresa = 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then Contador = Contador + TBAbrir.RecordCount
        TBAbrir.Close
        TBFornecedor!CPF_CNPJ = Contador + 1
End If

If cmbOrigem = "" Then
    TBFornecedor!idTipoEmpresa = Null
ElseIf cmbOrigem.Text = "Nacional" Then
        TBFornecedor!idTipoEmpresa = 1
    Else
        TBFornecedor!idTipoEmpresa = 0
End If
TBFornecedor!Estado = cmbuf.Text
TBFornecedor!cxpostal = txtcaixapostal.Text
TBFornecedor!CEP = IIf(txtCEP.Text = "", Null, txtCEP.Text)
TBFornecedor!Pessoa = cmbPessoa.Text
TBFornecedor!Email = IIf(txtEmail.Text = "", Null, LCase(txtEmail.Text))
TBFornecedor!Site = IIf(txtSite.Text = "", Null, LCase(txtSite.Text))
TBFornecedor!Fax = IIf(txtFax.Text = "", Null, txtFax.Text)
TBFornecedor!Tipo_endereco = cmbTipo_endereco
TBFornecedor!Tipo_bairro = cmbTipo_bairro
TBFornecedor!complemento = IIf(txtComplemento.Text = "", Null, txtComplemento.Text)
TBFornecedor!Pais = Txt_pais
If Txt_pais <> "" Then TBFornecedor!Codigo_pais = Txt_pais.ItemData(Txt_pais.ListIndex)
TBFornecedor!Categoria = txtCategoria.Text
TBFornecedor!RG_IM = txtIM_IE
TBFornecedor!Numero = txtNumero
If Chk_prospecto.Value = 1 Then TBFornecedor!Prospecto = True Else TBFornecedor!Prospecto = False
If Chk_enviar_NF.Value = 1 Then TBFornecedor!Enviar_NF = True Else TBFornecedor!Enviar_NF = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Fornecedores"
Direitos
ProcCarregaCombos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362F" Then
    If USMsgBox("Deseja realmente atualizar o código do país?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_fornecedores order by Pais", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir.MoveLast
            'PBLista.Min = 0
            'PBLista.Max = TBAbrir.RecordCount
            'PBLista.Value = 1
            Contador = 0
            TBAbrir.MoveFirst
            Do While TBAbrir.EOF = False
                If IsNull(TBAbrir!Pais) = True Or TBAbrir!Pais <> "" Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from Codigos_pais where Pais = '" & TBAbrir!Pais & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBAbrir!Codigo_pais = TBFI!CODIGO
                    Else
                        TBAbrir!Pais = ""
                        TBAbrir!Codigo_pais = Null
                    End If
                    TBAbrir.Update
                End If
                TBAbrir.MoveNext
                Contador = Contador + 1
                'PBLista.Value = Contador
            Loop
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Compras/Fornecedores"
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

Private Sub ProcAbrir()
On Error GoTo tratar_erro

frmcompras_fornecedores_localizar.Show 1

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
Troca_Duplicata = False
Financeiro_Contas_Recebidas = False
Engenharia_Normas = False
Qualidade_PPAP_PSW = False
Qualidade_PPAP_Plano = False
Qualidade_PPAP_FMEA = False
Qualidade_sistema = False
Engenharia = False
Compras_Fornecedores = True
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmCompras_fornecedores_RelUF.Show 1

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
With txtIDcliente
    .Text = ""
    .Locked = True
    .TabStop = False
End With
ProcLimpar
Novo_Fornecedor = True
Frame1.Enabled = True
cmbPessoa.SetFocus
ProcLimparTudo
USMsgBox "Após escolher o Tipo a origem o Pais a UF e digitar o CNPJ, clique no botão ao lado do CNPJ para buscar os dados do novo fornecedor na receita federal.", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame3.Enabled = False
Frame11.Enabled = False
Frame16.Enabled = False
Frame6.Enabled = False
Frame10.Enabled = False
Frame5.Enabled = False
procLimpacamposContatos
ProcLimpacamposEntrega
ProcLimpacamposCobranca
ProcLimpaFamilia
ProcLimpacampos_banco
ProcLimpaCampos_Segmento
ProcLimpaCampos_Outros
Novo_Fornecedor1 = False
Novo_Fornecedor2 = False
Novo_Fornecedor3 = False
Novo_Fornecedor4 = False
Novo_Fornecedor5 = False
Novo_Fornecedor6 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Fornecedor = True Then
    If USMsgBox("O fornecedor ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Fornecedor = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Fornecedor1 = True Then
    If USMsgBox("O contato ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_contato
        If Novo_Fornecedor1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Fornecedor4 = True Then
    If USMsgBox("A familia ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_familia
        If Novo_Fornecedor4 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Fornecedor5 = True Then
    If USMsgBox("O banco ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_comerciais
        If Novo_Fornecedor5 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Fornecedor = False
Novo_Fornecedor1 = False
Novo_Fornecedor4 = False
Novo_Fornecedor5 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro
    
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtStatus = "Liberado"
txtDtValidacao = ""
txtRespValidacao = ""
Chk_prospecto.Value = 0
Chk_enviar_NF.Value = 1
With cmbPessoa
    .ListIndex = -1
'    .Locked = False
'    .TabStop = True
End With
txtnomerazao.Text = ""
txtCategoria.ListIndex = -1
cmbOrigem.ListIndex = -1
txtcnpj.Text = "__.___.___/____-__"
txtCpf.Text = "___.___.___-__"
txtRG_IE.Text = ""
txtIM_IE.Text = ""
txtnomefantasia.Text = ""
Txt_pais.ListIndex = -1
cmbTipo_endereco.ListIndex = -1
txtendereco.Text = ""
txtNumero.Text = ""
cmbTipo_bairro.ListIndex = -1
txtBairro.Text = ""
cmbuf.ListIndex = -1
cmbCidade.ListIndex = -1
txtCidade.Text = ""
txtComplemento = ""
txtcaixapostal.Text = ""
txtCEP.Text = ""
txtTelefones.Text = ""
txtFax.Text = ""
txtEmail.Text = ""
txtSite.Text = ""
cmbRegimeTributario.ListIndex = -1
Cmb_tipo_transp.ListIndex = -1
cmbtransportadora.ListIndex = -1
ProcCarregaComboSetor Cmb_centro, "Setor is not null and DtBloq IS NULL  and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
Chk_nao_contribuinte_ICMS.Value = 0
CodigoLista = 0
Caption = "Administrativo - Compras - Fornecedores"
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtnomerazao.Text = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
Caption = "Administrativo - Compras - Fornecedores (Fornecedor : " & IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
txtData.Text = IIf(IsNull(TBFornecedor!DTCadastro), "", Format(TBFornecedor!DTCadastro, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBFornecedor!Responsavel), "", TBFornecedor!Responsavel)
txtDtValidacao.Text = IIf(IsNull(TBFornecedor!DtValidacao), "", TBFornecedor!DtValidacao)
txtRespValidacao.Text = IIf(IsNull(TBFornecedor!RespValidacao), "", TBFornecedor!RespValidacao)
With cmbPessoa
    If IsNull(TBFornecedor!Pessoa) = False And TBFornecedor!Pessoa <> "" Then
        .Text = TBFornecedor!Pessoa
'        .Locked = True
'        .TabStop = False
    Else
        .ListIndex = -1
'        .Locked = False
'        .TabStop = True
    End If
End With
If TBFornecedor!idTipoEmpresa = 1 Then
    If cmbPessoa = "JURÍDICA" Then
        If TBFornecedor!CPF_CNPJ <> "" Then
            CNPJ = Trim(TBFornecedor!CPF_CNPJ)
            With txtcnpj
                .PromptInclude = False
                .Text = CNPJ
                .PromptInclude = True
            End With
        End If
    Else
        If TBFornecedor!CPF_CNPJ <> "" Then
            CNPJ = Trim(TBFornecedor!CPF_CNPJ)
            With txtCpf
                .PromptInclude = False
                .Text = CNPJ
                .PromptInclude = True
            End With
        End If
    End If
End If
If TBFornecedor!idTipoEmpresa <> "" Then
    If TBFornecedor!idTipoEmpresa = 1 Then cmbOrigem = "Nacional"
    If TBFornecedor!idTipoEmpresa = 0 Then cmbOrigem = "Estrangeiro"
End If
If TBFornecedor!Presumido = True Then cmbRegimeTributario = "Lucro presumido"
If TBFornecedor!Simples = True Then cmbRegimeTributario = "Simples nacional"
If TBFornecedor!Real = True Then cmbRegimeTributario = "Lucro real"
If TBFornecedor!MEI = True Then cmbRegimeTributario = "MEI"
txtCEP.Text = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
txtRG_IE.Text = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
txtTelefones.Text = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
txtFax.Text = IIf(IsNull(TBFornecedor!Fax), "", TBFornecedor!Fax)
txtEmail.Text = IIf(IsNull(TBFornecedor!Email), "", TBFornecedor!Email)
txtSite.Text = IIf(IsNull(TBFornecedor!Site), "", TBFornecedor!Site)
txtnomefantasia.Text = IIf(IsNull(TBFornecedor!NomeFantasia), "", TBFornecedor!NomeFantasia)

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Usuarios_setor.* from Usuarios_setor where ID = " & IIf(IsNull(TBFornecedor!ID_CC), 0, TBFornecedor!ID_CC), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
        If IsNull(TBFI!DtBloq) = False Then
            Cmb_centro.AddItem TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
        End If
        Cmb_centro = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    Else
        If IsNull(TBFI!DtBloq) = False Then
            Cmb_centro.AddItem IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
        End If
        Cmb_centro = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    End If
End If
TBFI.Close

txtIM_IE = IIf(IsNull(TBFornecedor!RG_IM), "", TBFornecedor!RG_IM)
txtcaixapostal.Text = IIf(IsNull(TBFornecedor!cxpostal), "", TBFornecedor!cxpostal)
txtStatus.Text = IIf(IsNull(TBFornecedor!status), "", TBFornecedor!status)
txtNumero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
txtComplemento.Text = IIf(IsNull(TBFornecedor!complemento), "", TBFornecedor!complemento)
If TBFornecedor!Nao_contribuinte_ICMS = True Then Chk_nao_contribuinte_ICMS.Value = 1 Else Chk_nao_contribuinte_ICMS.Value = 0
   
NomeCampo = "o IQF"
If IsNull(TBFornecedor!Categoria) = False And TBFornecedor!Categoria <> "" Then txtCategoria.Text = TBFornecedor!Categoria
NomeCampo = "o tipo do endereço"
If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then cmbTipo_endereco.Text = TBFornecedor!Tipo_endereco
NomeCampo = "o tipo do bairro"
If IsNull(TBFornecedor!Tipo_bairro) = False And TBFornecedor!Tipo_bairro <> "" Then cmbTipo_bairro.Text = TBFornecedor!Tipo_bairro
NomeCampo = "o país"
If IsNull(TBFornecedor!Pais) = False And TBFornecedor!Pais <> "" Then Txt_pais = TBFornecedor!Pais
NomeCampo = "o estado"
If IsNull(TBFornecedor!Estado) = False And TBFornecedor!Estado <> "" Then cmbuf.Text = TBFornecedor!Estado
txtendereco.Text = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
txtBairro.Text = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
If TBFornecedor!Prospecto = True Then Chk_prospecto.Value = 1 Else Chk_prospecto.Value = 0
If TBFornecedor!Enviar_NF = True Then Chk_enviar_NF.Value = 1 Else Chk_enviar_NF.Value = 0
NomeCampo = "a Cidade"
If TBFornecedor!Estado <> "EX" Then
    cmbCidade.Visible = True
    txtCidade.Visible = False
    If IsNull(TBFornecedor!Cidade) = False And TBFornecedor!Cidade <> "" Then cmbCidade = TBFornecedor!Cidade
Else
    cmbCidade.Visible = False
    txtCidade.Visible = True
    txtCidade.Text = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
End If
NomeCampo = "o tipo da transportadora"
If IsNull(TBFornecedor!Tipo_transp) = False And TBFornecedor!Tipo_transp <> "" Then
    Select Case TBFornecedor!Tipo_transp
        Case "C": Cmb_tipo_transp = "Cliente"
        Case "F": Cmb_tipo_transp = "Fornecedor"
        Case "E": Cmb_tipo_transp = "Empresa"
    End Select
End If
NomeCampo = "a transportadora"
If IsNull(TBFornecedor!Transportadora) = False And TBFornecedor!Transportadora <> "" Then cmbtransportadora = TBFornecedor!Transportadora

2:
    Novo_Fornecedor = False
    With txtIDcliente
        .Locked = False
        .TabStop = True
    End With
    ProcLimparTudo

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desse fornecedor."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDados_Outros()
On Error GoTo tratar_erro

If IsNull(TBFornecedor!Obs) = False Then txtObs = TBFornecedor!Obs
If TBFornecedor!Fornecedor = "C" Then
    Chk_certificado.Value = 1
ElseIf TBFornecedor!Fornecedor = "A" Then
        Chk_avaliado.Value = 1
    ElseIf TBFornecedor!Fornecedor = "F" Then
            chkAprovada_Cliente.Value = 1
        Else
            Chk_avaliado.Value = 0
            Chk_certificado.Value = 0
            chkAprovada_Cliente.Value = 0
End If
txtdata_venc = IIf(IsNull(TBFornecedor!Data_venc), "__/__/____", Format(TBFornecedor!Data_venc, "dd/mm/yyyy"))
Txt_ICMS_ind = IIf(IsNull(TBFornecedor!ICMS_ind), 0, TBFornecedor!ICMS_ind)
txtCertificado = IIf(IsNull(TBFornecedor!Certificado), "", TBFornecedor!Certificado)
Txt_caminho_certificado = IIf(IsNull(TBFornecedor!Caminho_certificado), "", TBFornecedor!Caminho_certificado)
If TBFornecedor!Sedex = True Then chkSedex.Value = 1 Else chkSedex.Value = 0
If TBFornecedor!SimplesICMSST = True Then chkICMSST.Value = 1 Else chkICMSST.Value = 0
If TBFornecedor!Designado = True Then chkDesignado.Value = 1 Else chkDesignado.Value = 0
If IsNull(TBFornecedor!Banco) = False And TBFornecedor!Banco <> "" Then cmbBanco = TBFornecedor!Banco
If IsNull(TBFornecedor!Tipo_doc) = False And TBFornecedor!Tipo_doc <> "" Then cmbTipo_doc = TBFornecedor!Tipo_doc

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_banco_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_banco
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Compras_fornecedores", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
Proximo:
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_banco, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_banco_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_banco
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "dados bancário", "excluir estes", True, True) = False Then
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

Private Sub lista_banco_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_banco.ListItems.Count = 0 Then Exit Sub
ProcLimpacampos_banco
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores_banco where id = " & lista_banco.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtID_banco = TBLISTA!ID
    txtData5 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel5 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    txtBanco = IIf(IsNull(TBLISTA!Banco), "", TBLISTA!Banco)
    txtAgencia = IIf(IsNull(TBLISTA!Agencia), "", TBLISTA!Agencia)
    txtConta = IIf(IsNull(TBLISTA!Conta), "", TBLISTA!Conta)
    Novo_Fornecedor5 = False
    CodigoLista5 = lista_banco.SelectedItem.index
End If
TBLISTA.Close
Frame10.Enabled = True

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
                    If .ListItems(InitFor).ListSubItems(4) = "Sim" Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    
                    IDCliente = .ListItems(InitFor)
                    ProcVerificaRegistroUtilizadoSemMsg "Cotacao_fornecedor", "idforn = " & IDforn
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Compras_pedido", "idfornecedor = " & IDforn
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Compras_programa", "id_Forn = " & IDforn
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "CQ_RNC", "id_forn = " & IDforn
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "Projproduto_fornecedor", "Idfornecedor = " & IDforn
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_ContasPagar", "int_codforn = " & IDforn & " and txt_fornecedor = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal", "Id_Int_Cliente = " & IDCliente & " and txt_Razao_Nome = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
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

Private Sub Lista_contato_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_contato
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Compras_fornecedores", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_contato, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_contato_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_contato
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "contato", "excluir este", True, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_contato_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_contato.ListItems.Count = 0 Then Exit Sub
procLimpacamposContatos
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Contatos_fornecedor where IdContato = " & Lista_contato.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtIDContato.Text = TBLISTA!idcontato
    txtData1 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel1 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    txtNomeContato.Text = IIf(IsNull(TBLISTA!Nome), "", TBLISTA!Nome)
    txtdepartamento.Text = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
    txttelcontato.Text = IIf(IsNull(TBLISTA!ramal), "", TBLISTA!ramal)
    TxtEmail_Contato.Text = IIf(IsNull(TBLISTA!Email), "", TBLISTA!Email)
    If TBLISTA!Enviar_NFe = True Then Chk_enviar_NFe.Value = 1 Else Chk_enviar_NFe.Value = 0
    If TBLISTA!Enviar_boleto = True Then Chk_enviar_boleto.Value = 1 Else Chk_enviar_boleto.Value = 0
    Novo_Fornecedor1 = False
    CodigoLista1 = Lista_contato.SelectedItem.index
End If
TBLISTA.Close
Frame3.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_familia_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_familia
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Compras_fornecedores", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
Proximo:
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_familia, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_familia_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_familia
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "família", "excluir esta", True, True) = False Then
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

Private Sub lista_familia_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_familia.ListItems.Count = 0 Then Exit Sub
cmbfamilia.ListIndex = -1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores_familia where idFamilia = " & lista_familia.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtid_familia = TBLISTA!idFamilia
    txtData4 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel4 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    If IsNull(TBLISTA!Familia) = False And TBLISTA!Familia <> "" Then cmbfamilia = TBLISTA!Familia
1:
    Novo_Fornecedor4 = False
    CodigoLista4 = lista_familia.SelectedItem.index
End If
TBLISTA.Close
Frame6.Enabled = True

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a família deste fornecedor."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Cmb_opcao_lista = "Excluir" Then
                If .ListItems(InitFor).ListSubItems(4) = "Sim" Then
                    USMsgBox ("Não é permitido excluir este fornecedor, pois o mesmo está validado."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                                
                Mensagem = "Não é possível excluir este fornecedor, pois o mesmo está sendo utilizado no módulo"
                IDforn = .ListItems(InitFor)
                ProcVerificaRegistroUtilizado "Cotacao_fornecedor", "idforn = " & IDforn, "Compras/Cotação"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Compras_pedido", "idfornecedor = " & IDforn, "Compras/Pedido"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Compras_programa", "id_Forn = " & IDforn, "Compras/Programação"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "CQ_RNC", "id_forn = " & IDforn, "Qualidade/RNC"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Projproduto_fornecedor", "Idfornecedor = " & IDforn, "Engenharia/Produtos e serviços"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_ContasPagar", "int_codforn = " & IDforn & " and txt_fornecedor = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Financeiro/Contas a pagar"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal", "Id_Int_Cliente = " & IDforn & " and txt_Razao_Nome = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Faturamento/Nota fiscal/Terceiros"
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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Novo_Fornecedor = False
txtIDcliente = ""
txtIDcliente = Lista.SelectedItem
CodigoLista = Lista.SelectedItem.index
Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Segmento_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_Segmento
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Compras_fornecedores", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
Proximo:
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_Segmento, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Segmento_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_Segmento
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "segmento", "excluir este", True, True) = False Then
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

Private Sub Lista_Segmento_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_Segmento.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos_Segmento
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores_segmentos where id = " & Lista_Segmento.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtID_segmento = TBLISTA!ID
    txtData_segmento = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResp_segmento = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    txtSegmento = TBLISTA!Segmento
    Novo_Fornecedor6 = False
    CodigoLista6 = Lista_Segmento.SelectedItem.index
End If
TBLISTA.Close
Frame5.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listacobranca_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With listacobranca
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Compras_fornecedores", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal", "ID_cobranca = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView listacobranca, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listacobranca_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With listacobranca
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "endereço de cobrança", "excluir este", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Mensagem = "Não é permitido excluir este local de cobrança, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal_NFe", "ID_cobranca = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub listacobranca_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If listacobranca.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_cobranca where idcobranca = " & listacobranca.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtID_cobranca.Text = TBLISTA!idCobranca
    txtData3 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel3 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    If Left(cmbPessoa, 8) = "JURÍDICA" Then
        txtCNPJ_cobranca = IIf(IsNull(TBLISTA!CNPJ), "__.___.___/____-__", IIf(Len(TBLISTA!CNPJ) > 14, TBLISTA!CNPJ, "__.___.___/____-__"))
    Else
        txtCNPJ_cobranca = IIf(IsNull(TBLISTA!CNPJ), "___.___.___-__", IIf(Len(TBLISTA!CNPJ) > 14, "___.___.___-__", TBLISTA!CNPJ))
    End If
    txtendereco_cobranca.Text = IIf(IsNull(TBLISTA!endereco_Cobranca), "", TBLISTA!endereco_Cobranca)
    txtNumero_cobranca.Text = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
    txtComplemento_cobranca.Text = IIf(IsNull(TBLISTA!complemento), "", TBLISTA!complemento)
    txtbairro_cobranca.Text = IIf(IsNull(TBLISTA!bairro_Cobranca), "", TBLISTA!bairro_Cobranca)
    txtcidade_cobranca.Text = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
    txtcxpostal_cobranca.Text = IIf(IsNull(TBLISTA!cxpostal_Cobranca), "", TBLISTA!cxpostal_Cobranca)
    mskcep_cobranca = IIf(IsNull(TBLISTA!cep_Cobranca), "", TBLISTA!cep_Cobranca)
    txttel1_cobranca.Text = IIf(IsNull(TBLISTA!tel1_Cobranca), "", TBLISTA!tel1_Cobranca)
    txttel2_cobranca.Text = IIf(IsNull(TBLISTA!tel2_Cobranca), "", TBLISTA!tel2_Cobranca)
    txttel3_cobranca.Text = IIf(IsNull(TBLISTA!tel3_Cobranca), "", TBLISTA!tel3_Cobranca)
    txttel4_cobranca.Text = IIf(IsNull(TBLISTA!tel4_Cobranca), "", TBLISTA!tel4_Cobranca)
    txtfax_cobranca.Text = IIf(IsNull(TBLISTA!fax_Cobranca), "", TBLISTA!fax_Cobranca)
    txtemail_cobranca.Text = IIf(IsNull(TBLISTA!email_Cobranca), "", TBLISTA!email_Cobranca)
    txtSite_cobranca.Text = IIf(IsNull(TBLISTA!site_cobranca), "", TBLISTA!site_cobranca)
    NomeCampo = "o tipo do endereço"
    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then cmbTipo_endereco_cobranca.Text = TBLISTA!Tipo_endereco
    NomeCampo = "o tipo do bairro"
    If IsNull(TBLISTA!Tipo_bairro) = False And TBLISTA!Tipo_bairro <> "" Then cmbTipo_bairro_cobranca.Text = TBLISTA!Tipo_bairro
    NomeCampo = "o estado"
    If IsNull(TBLISTA!uf_Cobranca) = False And TBLISTA!uf_Cobranca <> "" Then cmbuf_cobranca = TBLISTA!uf_Cobranca
    NomeCampo = "a cidade"
    If TBLISTA!uf_Cobranca <> "EX" Then
        cmbCidade_cobranca.Visible = True
        txtcidade_cobranca.Visible = False
        cmbCidade_cobranca.Text = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
    Else
        cmbCidade_cobranca.Visible = False
        txtcidade_cobranca.Visible = True
        txtcidade_cobranca.Text = IIf(IsNull(TBLISTA!cidade_Cobranca), "", TBLISTA!cidade_Cobranca)
    End If
1:
    Frame16.Enabled = True
    CodigoLista3 = listacobranca.SelectedItem.index
    Novo_Fornecedor3 = False
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " do endereço de cobrança desse fornecedor."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEntrega_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEntrega
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Compras_fornecedores", "IDcliente = " & txtIDcliente, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "Compras_comercial", "ID_entrega = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Nota_Fiscal_NFe", "ID_entrega = " & .ListItems(InitFor)
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaEntrega, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEntrega_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEntrega
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "endereço de entrega", "excluir este", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Mensagem = "Não é permitido excluir este local para entrega, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "Compras_comercial", "ID_entrega = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Dados_Nota_Fiscal_NFe", "ID_entrega = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEntrega_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaEntrega.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from clientes_entrega where identrega = " & ListaEntrega.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtID_entrega.Text = TBLISTA!identrega
    txtData2 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel2 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    If Left(cmbPessoa, 8) = "JURÍDICA" Then
        txtCNPJ_entrega = IIf(IsNull(TBLISTA!CNPJ), "__.___.___/____-__", IIf(Len(TBLISTA!CNPJ) > 14, TBLISTA!CNPJ, "__.___.___/____-__"))
    Else
        txtCNPJ_entrega = IIf(IsNull(TBLISTA!CNPJ), "___.___.___-__", IIf(Len(TBLISTA!CNPJ) > 14, "___.___.___-__", TBLISTA!CNPJ))
    End If
    txtEndereco_entrega.Text = IIf(IsNull(TBLISTA!endereco_entrega), "", TBLISTA!endereco_entrega)
    txtNumero_entrega.Text = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
    txtComplemento_entrega.Text = IIf(IsNull(TBLISTA!complemento), "", TBLISTA!complemento)
    txtBairro_entrega.Text = IIf(IsNull(TBLISTA!bairro_entrega), "", TBLISTA!bairro_entrega)
    txtcxpostal_entrega.Text = IIf(IsNull(TBLISTA!cxpostal_entrega), "", TBLISTA!cxpostal_entrega)
    mskcep_entrega = IIf(IsNull(TBLISTA!cep_entrega), "", TBLISTA!cep_entrega)
    txttel1_entrega.Text = IIf(IsNull(TBLISTA!tel1_entrega), "", TBLISTA!tel1_entrega)
    txttel2_entrega.Text = IIf(IsNull(TBLISTA!tel2_entrega), "", TBLISTA!tel2_entrega)
    txttel3_entrega.Text = IIf(IsNull(TBLISTA!tel3_entrega), "", TBLISTA!tel3_entrega)
    txttel4_entrega.Text = IIf(IsNull(TBLISTA!tel4_entrega), "", TBLISTA!tel4_entrega)
    txtfax_entrega.Text = IIf(IsNull(TBLISTA!fax_entrega), "", TBLISTA!fax_entrega)
    txtemail_entrega.Text = IIf(IsNull(TBLISTA!email_entrega), "", TBLISTA!email_entrega)
    txtsite_entrega.Text = IIf(IsNull(TBLISTA!Site_Entrega), "", TBLISTA!Site_Entrega)
    NomeCampo = "o tipo do endereço"
    If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then cmbTipo_endereco_entrega.Text = TBLISTA!Tipo_endereco
    NomeCampo = "o tipo do bairro"
    If IsNull(TBLISTA!Tipo_bairro) = False And TBLISTA!Tipo_bairro <> "" Then cmbTipo_bairro_entrega.Text = TBLISTA!Tipo_bairro
    NomeCampo = "o estado"
    If IsNull(TBLISTA!uf_entrega) = False And TBLISTA!uf_entrega <> "" Then cmbuf_entrega = TBLISTA!uf_entrega
    NomeCampo = "a cidade"
    If TBLISTA!uf_entrega <> "EX" Then
        cmbCidade_Entrega.Visible = True
        txtCidade_Entrega.Visible = False
        cmbCidade_Entrega.Text = IIf(IsNull(TBLISTA!cidade_entrega), "", TBLISTA!cidade_entrega)
    Else
        cmbCidade_Entrega.Visible = False
        txtCidade_Entrega.Visible = True
        txtCidade_Entrega.Text = IIf(IsNull(TBLISTA!cidade_entrega), "", TBLISTA!cidade_entrega)
    End If
1:
    Frame11.Enabled = True
    CodigoLista2 = ListaEntrega.SelectedItem.index
    Novo_Fornecedor2 = False
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " do endereço de entrega desse fornecedor."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If

    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_avaliado_Click()
On Error GoTo tratar_erro

If Chk_avaliado.Value = 1 Then
    Chk_certificado.Value = 0
    chkAprovada_Cliente.Value = 0
    Frame7.Enabled = True
    Frame8.Enabled = True
Else
    If Chk_certificado.Value = 0 And chkAprovada_Cliente.Value = 0 Then
        Frame7.Enabled = False
        Frame8.Enabled = False
        txtdata_venc.Text = "__/__/____"
        txtCertificado = ""
        Txt_caminho_certificado = ""
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_certificado_Click()
On Error GoTo tratar_erro

If Chk_certificado.Value = 1 Then
    Chk_avaliado.Value = 0
    chkAprovada_Cliente.Value = 0
    Frame7.Enabled = True
    Frame8.Enabled = True
Else
    If Chk_avaliado.Value = 0 And chkAprovada_Cliente.Value = 0 Then
        Frame7.Enabled = False
        Frame8.Enabled = False
        txtdata_venc.Text = "__/__/____"
        txtCertificado = ""
        Txt_caminho_certificado = ""
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtIDcliente = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
USToolBar2.Visible = False
Select Case SSTab1.Tab
    Case 0:
        If Lista.Visible = True Then Lista.SetFocus
        With Lista
            .Visible = True
            .Top = Frame1.Top + Frame1.Height
            .Height = Frame2.Top - .Top
        End With
        Frame2.Visible = True
    Case 1:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista_contato.SetFocus
        ProcCarregaListaContatos
    Case 2:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ListaEntrega.SetFocus
        Proccarregalistaentrega
    Case 3:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        listacobranca.SetFocus
        ProcCarregalistacobranca
    Case 4:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_familia.SetFocus
        ProcCarregaListaFamilia
    Case 5:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_banco.SetFocus
        ProcCarregalista_banco
    Case 6:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame2.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista_Segmento.SetFocus
        ProcCarregaListaSegmento
    Case 7:
        With Lista
            .Visible = True
            .Top = Frame18.Top + Frame18.Height
            .Height = Frame2.Top - .Top
        End With
        Frame2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista.SetFocus
    Case 8:
        USToolBar2.Visible = True
        Lista.Visible = False
        Frame15.Visible = False
        'PBLista.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Cmb_empresa.SetFocus
        
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Fornecedor = True Then
    USMsgBox ("Salve o fornecedor antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaFamilia()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaFamilia()
On Error GoTo tratar_erro

txtData4 = Format(Date, "dd/mm/yy")
txtResponsavel4 = pubUsuario
txtid_familia = 0
cmbfamilia.ListIndex = -1
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Segmento()
On Error GoTo tratar_erro

txtData_segmento = Format(Date, "dd/mm/yy")
txtResp_segmento = pubUsuario
txtID_segmento = 0
txtSegmento = ""
CodigoLista6 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombos()
On Error GoTo tratar_erro

ProcCarregaFamilia
ProcCarregaComboPais Txt_pais
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboBancoFinanceiro cmbBanco, "txt_Descricao is not null", True
ProcCarregaComboTipoDocto cmbTipo_doc, "Tipo = 'P'"
With txtCategoria
    .Clear
    .AddItem "A"
    .AddItem "B"
    .AddItem "C"
    .AddItem "D"
    .AddItem "E"
    .AddItem "F"
    .AddItem "G"
    .AddItem "H"
    .AddItem "I"
    .AddItem "J"
End With
If txtIDcliente <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

Txt_pais.ListIndex = -1
cmbuf.ListIndex = -1
cmbCidade.ListIndex = -1
cmbtransportadora.ListIndex = -1
Cmb_centro.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_fornecedores where idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Categoria) = False And TBAbrir!Categoria <> "" Then txtCategoria = TBAbrir!Categoria
    If IsNull(TBAbrir!Pais) = False And TBAbrir!Pais <> "" Then Txt_pais = TBAbrir!Pais
    If IsNull(TBAbrir!Estado) = False And TBAbrir!Estado <> "" Then cmbuf = TBAbrir!Estado
    If IsNull(TBAbrir!Cidade) = False And TBAbrir!Cidade <> "" Then cmbCidade = TBAbrir!Cidade
    If IsNull(TBAbrir!Transportadora) = False And TBAbrir!Transportadora <> "" Then
        Select Case TBAbrir!Tipo_transp
            Case "C": Cmb_tipo_transp = "Cliente"
            Case "F": Cmb_tipo_transp = "Fornecedor"
            Case "E": Cmb_tipo_transp = "Empresa"
        End Select
        cmbtransportadora = TBAbrir!Transportadora
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Usuarios_setor.* from Usuarios_setor where ID = " & IIf(IsNull(TBAbrir!ID_CC), 0, TBAbrir!ID_CC), Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
            If IsNull(TBFI!DtBloq) = False Then
                Cmb_centro.AddItem TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
                Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
            End If
            Cmb_centro = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
        Else
            If IsNull(TBFI!DtBloq) = False Then
                Cmb_centro.AddItem IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
                Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
            End If
            Cmb_centro = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
        End If
    End If
    TBFI.Close
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaFamilia()
On Error GoTo tratar_erro

lista_familia.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from compras_fornecedores_familia where IDCliente = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and Tipo = 'F'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With lista_familia.ListItems
            .Add , , TBLISTA!idFamilia
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Familia), "", (TBLISTA!Familia))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaSegmento()
On Error GoTo tratar_erro

Lista_Segmento.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores_segmentos where IDFornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista_Segmento.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Segmento), "", (TBLISTA!Segmento))
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_banco()
On Error GoTo tratar_erro

lista_banco.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_fornecedores_banco where id_fornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and tipo = 'F' order by banco", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    'PBLista.Min = 0
    'PBLista.Max = TBLISTA.RecordCount
    'PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With lista_banco.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Banco), "", TBLISTA!Banco)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Agencia), "", TBLISTA!Agencia)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Conta), "", TBLISTA!Conta)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        'PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpacampos_banco()
On Error GoTo tratar_erro

txtData5 = Format(Date, "dd/mm/yy")
txtResponsavel5 = pubUsuario
txtBanco = ""
txtAgencia = ""
txtConta = ""
txtID_banco = 0
CodigoLista5 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_Outros()
On Error GoTo tratar_erro

Chk_certificado.Value = 0
Chk_avaliado.Value = 0
chkAprovada_Cliente.Value = 0
txtdata_venc = "__/__/____"
txtCertificado = ""
Txt_caminho_certificado = ""

Txt_ICMS_ind = ""
txtObs.Text = ""
cmbBanco.ListIndex = -1
cmbTipo_doc.ListIndex = -1
chkSedex.Value = 0
chkICMSST.Value = 0
chkDesignado.Value = 0
CodigoLista6 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_ICMS_ind_Change()
On Error GoTo tratar_erro

If Txt_ICMS_ind.Text <> "" Then
    VerifNumero = Txt_ICMS_ind.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ICMS_ind.Text = ""
        Txt_ICMS_ind.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtemail_cobranca_LostFocus()
On Error GoTo tratar_erro

txtemail_cobranca = LCase(txtemail_cobranca)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TxtEmail_Contato_LostFocus()
On Error GoTo tratar_erro

TxtEmail_Contato = LCase(TxtEmail_Contato)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtemail_entrega_LostFocus()
On Error GoTo tratar_erro

txtemail_entrega = LCase(txtemail_entrega)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEmail_LostFocus()
On Error GoTo tratar_erro

txtEmail = LCase(txtEmail)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDcliente_Change()
On Error GoTo tratar_erro

If Novo_Fornecedor = False Then
    ProcLimpar
    ProcLimpaCampos_Outros
End If
If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    If Novo_Fornecedor = False Then
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & IIf(txtIDcliente = "", 0, txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then
            ProcPuxaDados
            procPuxadados_Comerciais
            ProcPuxaDados_Outros
            Frame1.Enabled = True
        Else
            Frame1.Enabled = False
        End If
        TBFornecedor.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDCliente_LostFocus()
On Error GoTo tratar_erro


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

Private Sub txtSite_cobranca_LostFocus()
On Error GoTo tratar_erro

txtSite_cobranca = LCase(txtSite_cobranca)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtsite_entrega_LostFocus()
On Error GoTo tratar_erro

txtsite_entrega = LCase(txtsite_entrega)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtSite_LostFocus()
On Error GoTo tratar_erro

Txt_site = LCase(Txt_site)

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
    Case 8: procFiltrar_todos
    Case 9: ProcStatus
    Case 10: ProcValidarRegistros Lista, "Compras/Fornecedores"
    Case 11: procAtualiza
    Case 13: ProcAjuda
    Case 14: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoTab
    Case 2: ProcSalvarTab
    Case 3: ProcExcluirTab
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

Sub ProcNovoTab()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 1: If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "contato", "criar novo", True, True) = False Then Exit Sub
            procNovo_Contato
    Case 2: If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "endereço de entrega", "criar novo", True, True) = False Then Exit Sub
            procNovo_entrega
    Case 3: If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "endereço de cobrança", "criar novo", True, True) = False Then Exit Sub
            procNovo_cobranca
    Case 4: If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "família", "criar novo", True, True) = False Then Exit Sub
            procNovo_familia
    Case 5: If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "dados bancários", "criar novos", True, True) = False Then Exit Sub
            ProcNovo_comerciais
    Case 6: If FunVerificaRegistroValidado("Compras_fornecedores", "IDcliente = " & txtIDcliente, "fornecedor", "segmento", "criar novo", True, True) = False Then Exit Sub
            procNovo_segmento
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarTab()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 1: procSalvar_contato
    Case 2: procSalvar_entrega
    Case 3: procSalvar_cobranca
    Case 4: procSalvar_familia
    Case 5: procSalvar_comerciais
    Case 6: procSalvar_segmento
    Case 7: ProcSalvar
    Case 8: ProcSalvar_comercial
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadadosComercial()
On Error GoTo tratar_erro

TBProduto!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBProduto!condicoes = IIf(txtCondicoes.Text = "", Null, txtCondicoes.Text)
'TBProduto!calculos = txtcalculos.Text
TBProduto!transporte = IIf(txttransporte.Text = "", Null, txttransporte.Text)
TBProduto!impostos = IIf(txtimpostos.Text = "", Null, txtimpostos.Text)
'TBProduto!garantia = txtgarantia.Text
'TBProduto!reajuste = txtReajuste.Text
TBProduto!validade = IIf(txtValidade.Text = "", Null, txtValidade.Text)
TBProduto!IDCFOP = IIf(txtID_cfop = "", Null, txtID_cfop)
TBProduto!CFOP = txtCFOP
TBProduto!descricaoCFOP = IIf(txtOperacao = "", Null, txtOperacao.Text)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposComerciais()
On Error GoTo tratar_erro

txtcalculos.Text = "N/A"
txtimpostos.Text = "N/A"
txtCondicoes.Text = "N/A"
txtgarantia.Text = "N/A"
txtReajuste.Text = "N/A"
txttransporte.Text = "N/A"
txtValidade.Text = "N/A"
txtID_cfop.Text = ""
txtCFOP.Text = ""
txtOperacao = ""
CodigoLista7 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxadados_Comerciais()
On Error GoTo tratar_erro


Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM Compras_fornecedores_DadosComerciais WHERE IDfornecedor = " & IIf(txtIDcliente = "", 0, txtIDcliente) & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
'    txtcalculos.Text = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
    txtimpostos.Text = IIf(IsNull(TBCotacao!impostos), "", TBCotacao!impostos)
    txtCondicoes.Text = IIf(IsNull(TBCotacao!condicoes), "", TBCotacao!condicoes)
'    txtgarantia.Text = IIf(IsNull(TBCotacao!garantia), "", TBCotacao!garantia)
'    txtReajuste.Text = IIf(IsNull(TBCotacao!reajuste), "", TBCotacao!reajuste)
    txttransporte.Text = IIf(IsNull(TBCotacao!transporte), "", TBCotacao!transporte)
    txtValidade.Text = IIf(IsNull(TBCotacao!validade), "", TBCotacao!validade)
    txtID_cfop = IIf(IsNull(TBCotacao!IDCFOP), "", TBCotacao!IDCFOP)
    txtCFOP = IIf(IsNull(TBCotacao!CFOP), "", TBCotacao!CFOP)
    txtOperacao = IIf(IsNull(TBCotacao!descricaoCFOP), "", TBCotacao!descricaoCFOP)
Else
    txtimpostos.Text = ""
    txtCondicoes.Text = ""
    txttransporte.Text = ""
    txtValidade.Text = ""
    txtID_cfop = ""
    txtCFOP = ""
    txtOperacao = ""
End If
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcSalvar_comercial()
On Error GoTo tratar_erro
  
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * FROM Compras_fornecedores_DadosComerciais WHERE IDfornecedor = " & txtIDcliente.Text & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
   ' If FunVerificaRegistroValidado("Fornecedores", "IDcliente = " & txtIDCliente, "fornecedor", "estes dados comerciais", "alterar", True, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dados comerciais"
Else
    TBProduto.AddNew
    USMsgBox ("Dados comerciais do fornecedor cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo dados comerciais"
    TBProduto!IDFornecedor = txtIDcliente
End If
ProcEnviadadosComercial
TBProduto.Update
ID_documento = TBProduto!ID
TBProduto.Close
'==================================
Modulo = "Compras/Fornecedores"
Documento = "Fornecedor: " & txtnomerazao & " - Empresa: " & Cmb_empresa
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
Sub ProcExcluirTab()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 1: procExcluir_contato
    Case 2: procExcluir_entrega
    Case 3: procExcluir_cobranca
    Case 4: procExcluir_familia
    Case 5: ProcExcluir_comercial
    Case 6: procExcluir_Segmento
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar_outros
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
