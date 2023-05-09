VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_programacao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Programação"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
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
      TabIndex        =   65
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
   Begin VB.ComboBox Cmb_empresa 
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
      ItemData        =   "frmCompras_programacao.frx":0000
      Left            =   260
      List            =   "frmCompras_programacao.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1695
      Width           =   6645
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17701
      _Version        =   393216
      Tab             =   2
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
      TabPicture(0)   =   "frmCompras_programacao.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Framelista"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtID"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "USImageList1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Produtos/Serviços"
      TabPicture(1)   =   "frmCompras_programacao.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USImageList2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtID_item"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lista_item"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "USToolBar2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Programação"
      TabPicture(2)   =   "frmCompras_programacao.frx":003C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtID_prog"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "USImageList3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   70
         Top             =   9090
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
            Left            =   3780
            TabIndex        =   11
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
            TabIndex        =   12
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   16
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_programacao.frx":0058
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
            TabIndex        =   15
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_programacao.frx":37FF
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
            TabIndex        =   13
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
            TabIndex        =   14
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_programacao.frx":7308
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
            TabIndex        =   17
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_programacao.frx":B3FA
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
            Left            =   4410
            TabIndex        =   75
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label32 
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   240
            Width           =   1095
         End
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   12870
         Top             =   510
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_programacao.frx":EC88
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -63150
         Top             =   510
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_programacao.frx":1406C
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   -63480
         Top             =   420
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCompras_programacao.frx":19A97
         Count           =   1
      End
      Begin VB.TextBox txtID_item 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73530
         TabIndex        =   59
         Text            =   "0"
         Top             =   3120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72750
         TabIndex        =   58
         Text            =   "0"
         Top             =   3780
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtID_prog 
         Height          =   285
         Left            =   3780
         TabIndex        =   57
         Text            =   "0"
         Top             =   4080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Notas fiscais"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7245
         Left            =   10155
         TabIndex        =   55
         Top             =   2760
         Width           =   5115
         Begin VB.TextBox txttotal 
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
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade total."
            Top             =   6810
            Width           =   2775
         End
         Begin MSComctlLib.ListView lista_nota 
            Height          =   5945
            Left            =   180
            TabIndex        =   34
            Top             =   315
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   10478
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
               Text            =   "Id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Nota fiscal"
               Object.Width           =   3219
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1941
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Quantidade"
               Object.Width           =   2646
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBlista3 
            Height          =   255
            Left            =   180
            TabIndex        =   69
            Top             =   6270
            Width           =   4725
            _ExtentX        =   8334
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
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total recebido"
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
            Left            =   2025
            TabIndex        =   61
            Top             =   6600
            Width           =   1020
         End
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
         Height          =   7245
         Left            =   75
         TabIndex        =   48
         Top             =   2760
         Width           =   10065
         Begin MSComctlLib.ListView lista_prog 
            Height          =   6675
            Left            =   150
            TabIndex        =   33
            Top             =   180
            Width           =   9705
            _ExtentX        =   17119
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Inicio prazo"
               Object.Width           =   2002
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Final prazo"
               Object.Width           =   2002
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   2178
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Qtde. a receber"
               Object.Width           =   2531
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Confirm."
               Object.Width           =   1826
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   3589
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Pedido"
               Object.Width           =   1764
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBLista2 
            Height          =   255
            Left            =   150
            TabIndex        =   68
            Top             =   6870
            Width           =   9735
            _ExtentX        =   17171
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   855
         Left            =   -74925
         TabIndex        =   40
         Top             =   1305
         Width           =   15195
         Begin VB.CommandButton cmdfiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1950
            Picture         =   "frmCompras_programacao.frx":2186F
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Filtrar por código interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtStatus_item 
            Alignment       =   2  'Center
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
            Left            =   12060
            TabIndex        =   23
            Top             =   390
            Width           =   2925
         End
         Begin VB.CommandButton cmdLocalizar_item 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2280
            Picture         =   "frmCompras_programacao.frx":21C8A
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Localizar produto (F2)"
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtDescricao 
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
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do item."
            Top             =   390
            Width           =   9375
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
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Código do item."
            Top             =   390
            Width           =   1755
         End
         Begin VB.Label Label27 
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
            Left            =   13245
            TabIndex        =   54
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
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
            Left            =   7012
            TabIndex        =   42
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
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
            Left            =   442
            TabIndex        =   41
            Top             =   180
            Width           =   1230
         End
      End
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1485
         Left            =   -74925
         TabIndex        =   36
         Top             =   1305
         Width           =   15195
         Begin VB.TextBox txtData_rev 
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
            Left            =   13770
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Data da revisão."
            Top             =   390
            Width           =   1215
         End
         Begin VB.TextBox txtRev 
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
            Left            =   13225
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Número de revisão da programação."
            Top             =   390
            Width           =   525
         End
         Begin VB.TextBox txtID_forn 
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
            Left            =   2310
            TabIndex        =   7
            ToolTipText     =   "Código do fornecedor."
            Top             =   1020
            Width           =   795
         End
         Begin VB.TextBox txtStatus 
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
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   1020
            Width           =   2115
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
            Left            =   8065
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3660
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
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1215
         End
         Begin VB.CommandButton cmdpesquisar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frmCompras_programacao.frx":21D8C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Localizar fornecedor."
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox txtPrograma 
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
            Left            =   11740
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Numero do programa."
            Top             =   390
            Width           =   1465
         End
         Begin VB.TextBox txtFornecedor 
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
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Fornecedor."
            Top             =   1020
            Width           =   11535
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3210
            TabIndex        =   64
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data rev."
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
            Left            =   13987
            TabIndex        =   63
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rev."
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
            Left            =   13300
            TabIndex        =   62
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label26 
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
            Left            =   960
            TabIndex        =   53
            Top             =   810
            Width           =   555
         End
         Begin VB.Label Label25 
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
            Left            =   9438
            TabIndex        =   52
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº programa"
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
            Left            =   11932
            TabIndex        =   39
            Top             =   180
            Width           =   1080
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
            Left            =   7275
            TabIndex        =   38
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Left            =   8475
            TabIndex        =   37
            Top             =   810
            Width           =   825
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1455
         Left            =   75
         TabIndex        =   43
         Top             =   1305
         Width           =   15195
         Begin DrawSuite2022.USCheckBox optFirme 
            Height          =   315
            Left            =   13290
            TabIndex        =   76
            Top             =   1020
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "Compra confirmada"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin VB.TextBox Txt_un_com 
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   990
            Width           =   915
         End
         Begin VB.TextBox Txt_un 
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
            Left            =   5550
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Unidade de estoque."
            Top             =   990
            Width           =   915
         End
         Begin VB.TextBox txtStatus_prog 
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
            Left            =   7410
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Status da programação."
            Top             =   990
            Width           =   5775
         End
         Begin VB.TextBox txtDescricao2 
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
            Left            =   1860
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do item."
            Top             =   390
            Width           =   13125
         End
         Begin VB.TextBox txtQtd 
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
            Left            =   3270
            TabIndex        =   29
            ToolTipText     =   "Quantidade."
            Top             =   990
            Width           =   2265
         End
         Begin VB.TextBox txtCodigo2 
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
            MaxLength       =   50
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Código do item."
            Top             =   390
            Width           =   1665
         End
         Begin MSComCtl2.DTPicker txtData_inicio 
            Height          =   315
            Left            =   180
            TabIndex        =   27
            ToolTipText     =   "Inicio do prazo."
            Top             =   990
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   172621825
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtData_fim 
            Height          =   315
            Left            =   1950
            TabIndex        =   28
            ToolTipText     =   "Final do prazo."
            Top             =   990
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   172687361
            CurrentDate     =   39057
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. com."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6615
            TabIndex        =   74
            Top             =   780
            Width           =   645
         End
         Begin VB.Label Label28 
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
            Height          =   195
            Left            =   10020
            TabIndex        =   56
            Top             =   780
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un. est."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5715
            TabIndex        =   51
            Top             =   780
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final do prazo"
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
            Left            =   2100
            TabIndex        =   50
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio do prazo"
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
            Left            =   322
            TabIndex        =   49
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. interno"
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
            Left            =   502
            TabIndex        =   47
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
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
            Height          =   180
            Left            =   8077
            TabIndex        =   46
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade"
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
            Left            =   3982
            TabIndex        =   45
            Top             =   780
            Width           =   840
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "á"
            Height          =   195
            Left            =   1680
            TabIndex        =   44
            Top             =   1050
            Width           =   90
         End
      End
      Begin MSComctlLib.ListView lista_item 
         Height          =   7530
         Left            =   -74925
         TabIndex        =   24
         Top             =   2175
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13282
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
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   18177
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lista 
         Height          =   6270
         Left            =   -74925
         TabIndex        =   10
         Top             =   2805
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11060
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
         NumItems        =   8
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
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Nº programa"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Data rev."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   8828
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   66
         Top             =   330
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
         ButtonCaption8  =   "Revisar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Revisar (F7)"
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
         ButtonWidth8    =   44
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Emitir PC"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Emitir PC das programação(ões) no período (F8)"
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
         ButtonLeft9     =   352
         ButtonTop9      =   2
         ButtonWidth9    =   50
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
         ButtonLeft10    =   404
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
         ButtonLeft11    =   456
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   460
         ButtonTop12     =   2
         ButtonWidth12   =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   498
         ButtonTop13     =   2
         ButtonWidth13   =   26
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   526
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   67
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
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
         ButtonCaption7  =   "Emitir PC"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Emitir PC das programação(ões) do produto/item no período (F8)"
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
         ButtonLeft7     =   268
         ButtonTop7      =   2
         ButtonWidth7    =   50
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
         ButtonLeft8     =   320
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   324
         ButtonTop9      =   2
         ButtonWidth9    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   362
         ButtonTop10     =   2
         ButtonWidth10   =   26
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
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
         ButtonState11   =   5
         ButtonLeft11    =   390
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   75
         TabIndex        =   20
         ToolTipText     =   "Unidade de estoque."
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
         ButtonToolTipText5=   "Produto anterior."
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
         ButtonToolTipText6=   "Próximo produto."
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
            Name            =   "Tahoma"
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
      End
   End
End
Attribute VB_Name = "frmCompras_programacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Programacao_Compras              As Boolean 'OK
Dim Novo_Programacao_Compras1             As Boolean 'OK
Dim Novo_Programacao_Compras2             As Boolean 'OK
Public Sql_Programacao_Compras_Localizar  As String 'OK
Dim TBLISTA_Compras_Programacao   As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=lLQHL9Qur68&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=40&feature=plcp")

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

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_programa order by Programatexto", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID = " & txtId)
    TBCompras.MovePrevious
    If TBCompras.BOF = False Then
        txtId = TBCompras!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCampos_Item
        ProcLimpaCampos_Prog
        ProcCarregaDados
        ProcCarregaLista_Item
        ProcCarregalista_Prog
    Else
        USMsgBox ("Fim dos cadastros de programação de compra."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Compras1 = False
Novo_Programacao_Compras2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior_prog()
On Error GoTo tratar_erro

If txtID_item = "" Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_programa_item where ID = " & txtId & " and codigo <> 'Null' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID_item = " & txtID_item)
    TBCompras.MovePrevious
    If TBCompras.BOF = False Then
        ProcLimpaCampos_Prog
        txtID_item = TBCompras!Id_Item
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_programa_item where ID_Item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
        txtCodigo = TBAbrir!CODIGO
        txtCodigo2 = TBAbrir!CODIGO
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select descricao from projproduto where desenho = '" & TBAbrir!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            txtdescricao = TBProduto!Descricao
            txtDescricao2 = TBProduto!Descricao
        End If
        TBProduto.Close
        ProcCarregalista_Prog
        txtTotal.Text = "0,0000"
        Proccarreganota
    Else
        USMsgBox ("Fim dos cadastros de produtos/serviços."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Compras2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_prog()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_prog
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) programação(ões) do produto/serviço?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_programacao where ID_prog = " & .ListItems(InitFor)
            ProcVerifExcluirPedido "Compras_pedido.ID_programa = " & txtId & " and compras_pedido_lista.ID_programacao = " & .ListItems(InitFor)
            '==================================
            Modulo = "Compras/Programação"
            Evento = "Excluir programação do produto/serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev & " - Cód. interno: " & txtCodigo & " - Descrição: " & txtdescricao
            Documento1 = "Data: " & .ListItems(InitFor).SubItems(1) & "-" & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) programação(ões) do produto/serviço antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Programação(ões) do produto/serviço excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Prog
    ProcCarregalista_Prog
    Frame7.Enabled = False
    optFirme.Enabled = True
    ProcAltera_Status txtID_item, txtId
    Novo_Programacao_Compras2 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_item()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With lista_item
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s) ?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from compras_programa_item where id_item = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Compras_programacao where id_item = " & .ListItems(InitFor)
            ProcVerifExcluirPedido "Compras_pedido.ID_programa = " & txtId & " and Compras_pedido_lista.Desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
            '==================================
            Modulo = "Compras/Programação"
            Evento = "Excluir produto/serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1) & " - Descrição: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produtos(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produtos(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Item
    ProcCarregaLista_Item
    Frame4.Enabled = False
    ProcAltera_Status txtID_item, txtId
    ProcLimparTudo
    Novo_Programacao_Compras1 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

If txtCodigo <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select desenho, descricao from projproduto where desenho like '" & txtCodigo & "%' and Compras = 'True' and Bloqueado = 'False' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtCodigo = TBProduto!Desenho
        txtdescricao = TBProduto!Descricao
    End If
    TBProduto.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_item_Click()
On Error GoTo tratar_erro

frmCompras_programacao_localizar_prod.Show 1

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
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido criar um novo produto/serviço, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos_Item
Novo_Programacao_Compras1 = True
Frame4.Enabled = True
txtCodigo.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_prog()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido criar uma nova programação, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCampos_Prog
Novo_Programacao_Compras2 = True
Frame7.Enabled = True
optFirme.Enabled = True
txtData_inicio.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Compras_Programacao.AbsolutePage <> 2 Then
    If TBLISTA_Compras_Programacao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Compras_Programacao.PageCount - 1)
    Else
        TBLISTA_Compras_Programacao.AbsolutePage = TBLISTA_Compras_Programacao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Compras_Programacao.AbsolutePage)
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
    TBLISTA_Compras_Programacao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Compras_Programacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Compras_Programacao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Compras_Programacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Compras_Programacao.AbsolutePage <> -3 Then
    If TBLISTA_Compras_Programacao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Compras_Programacao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Compras_Programacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Compras_Programacao.AbsolutePage = TBLISTA_Compras_Programacao.PageCount
ProcExibePagina (TBLISTA_Compras_Programacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpesquisar_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocForn False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = 0 Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_programa order by Programatexto", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID = " & txtId)
    TBCompras.MoveNext
    If TBCompras.EOF = False Then
        txtId = TBCompras!ID
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCampos_Item
        ProcLimpaCampos_Prog
        ProcCarregaDados
        ProcCarregaLista_Item
        ProcCarregalista_Prog
    Else
        USMsgBox ("Fim dos cadastros de programação de compra."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Compras1 = False
Novo_Programacao_Compras2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo_prog()
On Error GoTo tratar_erro

If txtID_item = "" Then Exit Sub
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_programa_item where ID = " & txtId & " and codigo <> 'Null' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.BOF = False Then
    TBCompras.Find ("ID_item = " & txtID_item)
    TBCompras.MoveNext
    If TBCompras.EOF = False Then
        ProcLimpaCampos_Prog
        txtID_item = TBCompras!Id_Item
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_programa_item where ID_Item = " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
        txtCodigo = TBAbrir!CODIGO
        txtCodigo2 = TBAbrir!CODIGO
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select descricao from projproduto where desenho = '" & TBAbrir!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            txtdescricao = TBProduto!Descricao
            txtDescricao2 = TBProduto!Descricao
        End If
        TBProduto.Close
        ProcCarregalista_Prog
        txtTotal.Text = "0,0000"
        Proccarreganota
    Else
        USMsgBox ("Fim dos cadastros de produtos/serviços."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Programacao_Compras2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRevisar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "criar revisão"
If txtPrograma = "" Then
    NomeCampo = "o programa"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Produto = True Then
    USMsgBox ("Salve a programação antes de cadastrar as revisões."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente criar uma revisão deste programa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    '==================================
    Modulo = "Compras/Programação"
    ID_documento = txtId
    Evento = "Revisar"
    Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from Compras_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    txtRev = IIf(IsNull(TBCotacao!Rev), 0, TBCotacao!Rev) + 1
    txtData_rev = Format(Date, "dd/mm/yy")
    IDlista = TBCotacao!programa
    IDAntigo = txtId
    TBCotacao.AddNew
    TBCotacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBCotacao!programa = IDlista
    TBCotacao!programatexto = txtPrograma
    TBCotacao!Data = Date
    TBCotacao!Responsavel = pubUsuario
    TBCotacao!Rev = txtRev
    TBCotacao!status = "ABERTO"
    TBCotacao!ID_forn = txtID_forn
    TBCotacao!via = "0"
    TBCotacao.Update
    txtId = TBCotacao!ID
    TBCotacao.Close
    
    Conexao.Execute "UPDATE Compras_programa Set Status = 'REVISADA', Data_rev = '" & Format(Date, "Short Date") & "' where ID = " & IDAntigo
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Compras_programa_item where ID = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Do While TBProduto.EOF = False
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Compras_programa_item", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!CODIGO = TBProduto!CODIGO
            TBGravar!ID = txtId
            TBGravar!Status_Item = "ABERTO"
            TBGravar.Update
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Compras_programacao where ID_item = " & TBProduto!Id_Item & " and quantidade > Qtderecebida", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Set TBPrograma = CreateObject("adodb.recordset")
                    TBPrograma.Open "Select * from Compras_programacao", Conexao, adOpenKeyset, adLockOptimistic
                    TBPrograma.AddNew
                    TBPrograma!ID = txtId
                    TBPrograma!Id_Item = TBGravar!Id_Item
                    TBPrograma!Un = TBAbrir!Un
                    TBPrograma!Unidade_com = TBAbrir!Unidade_com
                    TBPrograma!Data_inicio = TBAbrir!Data_inicio
                    TBPrograma!Data_fim = TBAbrir!Data_fim
                    TBPrograma!quantidade = TBAbrir!quantidade - TBAbrir!Qtderecebida
                    TBPrograma!Firme = False
                    TBPrograma!Status_prog = "PREVISÃO FUTURA"
                    TBPrograma!Ordenar = 3
                    TBPrograma.Update
                    TBPrograma.Close
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            
            TBGravar.Close
            TBProduto.MoveNext
        Loop
    End If
    TBProduto.Close
    
    USMsgBox ("Revisão do programa criada com sucesso."), vbInformation, "CAPRIND v5.0"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_programa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then ProcCarregaDados
    ProcCarregaLista_Item
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmitirPC()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "firmar a(s) programação(ões)"
If txtPrograma = "" Then
    NomeCampo = "o programa"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Programacao_Compras = True Then
    USMsgBox ("Salve o programa antes de emitir o(s) pedido(s) da(s) programação(ões)."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente emitir o(s) pedido(s) da(s) programação(ões) deste programa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
    DataTexto = InputBox("Favor informar o prazo final da(s) programação(ões).")
    If DataTexto = "" Then Exit Sub
    If IsDate(DataTexto) = False Then
        USMsgBox ("Esta data não é válida."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    DataFim = DataTexto
    
    Set TBPrograma = CreateObject("adodb.recordset")
    TBPrograma.Open "Select Compras_programa_item.codigo, Compras_programacao.*, projproduto.Descricao from (Compras_programa_item INNER JOIN Compras_programacao ON Compras_programa_item.Id_item = Compras_programacao.Id_item) INNER JOIN projproduto ON projproduto.Desenho = Compras_programa_item.Codigo where Compras_programacao.ID = " & txtId & " and Compras_programacao.Data_fim <= '" & DataFim & "' and Compras_programacao.Firme = 'False' order by Compras_programa_item.ID_item, Compras_programacao.ID_prog", Conexao, adOpenKeyset, adLockOptimistic
    If TBPrograma.EOF = False Then
        Do While TBPrograma.EOF = False
            TBPrograma!Firme = True
            TBPrograma!Status_prog = "ABERTO"
            TBPrograma!Ordenar = 2
            TBPrograma.Update
            ProcAlteraStatusProgramacao TBPrograma!ID_prog, TBPrograma!quantidade
            ProcAltera_Status TBPrograma!Id_Item, txtId
            
            ProcGerarPedido TBPrograma!ID_prog, TBPrograma!Data_fim, TBPrograma!CODIGO, TBPrograma!Descricao, TBPrograma!quantidade
            
            '==================================
            Modulo = "Compras/Programação"
            ID_documento = txtId
            Evento = "Firmar programações"
            Documento = "Nº programa: " & txtPrograma & " - Nº revisão: " & txtRev & " - Cód. interno: " & TBPrograma!CODIGO & " - Descrição: " & TBPrograma!Descricao
            Documento1 = "Data: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & "-" & Format(TBPrograma!Data_fim, "dd/mm/yy")
            ProcGravaEvento
            '==================================
            
            TBPrograma.MoveNext
        Loop
        USMsgBox ("Pedido(s) emitido(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Não foi encontrada nenhuma programação com o prazo final menor ou igual a " & Format(DataFim, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
    End If
    TBPrograma.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmitirPC_item()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "firmar a(s) programação(ões)"
If txtdescricao = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Programacao_Compras1 = True Then
    USMsgBox ("Salve o produto/serviçoantes de emitir o(s) pedido(s) da(s) programação(ões)."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente emitir o(s) pedido(s) da(s) programação(ões) deste produto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem:
    DataTexto = InputBox("Favor informar o prazo final da(s) programação(ões).")
    If DataTexto = "" Then Exit Sub
    If IsDate(DataTexto) = False Then
        USMsgBox ("Esta data não é válida."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    DataFim = DataTexto
    
    Set TBPrograma = CreateObject("adodb.recordset")
    TBPrograma.Open "Select Compras_programa_item.codigo, Compras_programacao.*, projproduto.Descricao from (Compras_programa_item INNER JOIN Compras_programacao ON Compras_programa_item.Id_item = Compras_programacao.Id_item) INNER JOIN projproduto ON projproduto.Desenho = Compras_programa_item.Codigo where Compras_programacao.ID = " & txtId & " and Compras_programa_item.ID_item = " & txtID_item & " and Compras_programacao.Data_fim <= '" & DataFim & "' and Compras_programacao.Firme = 'False' order by Compras_programa_item.ID_item, Compras_programacao.ID_prog", Conexao, adOpenKeyset, adLockOptimistic
    If TBPrograma.EOF = False Then
        Do While TBPrograma.EOF = False
            TBPrograma!Firme = True
            TBPrograma!Status_prog = "ABERTO"
            TBPrograma!Ordenar = 2
            TBPrograma.Update
            ProcAlteraStatusProgramacao TBPrograma!ID_prog, TBPrograma!quantidade
            ProcAltera_Status TBPrograma!Id_Item, txtId
            
            ProcGerarPedido TBPrograma!ID_prog, TBPrograma!Data_fim, TBPrograma!CODIGO, TBPrograma!Descricao, TBPrograma!quantidade
            
            '==================================
            Modulo = "Compras/Programação"
            ID_documento = txtID_item
            Evento = "Firmar programações do produto"
            Documento = "Nº programa: " & txtPrograma & " - Nº revisão: " & txtRev & " - Cód. interno: " & TBPrograma!CODIGO & " - Descrição: " & TBPrograma!Descricao
            Documento1 = "Data: " & Format(TBPrograma!Data_inicio, "dd/mm/yy") & "-" & Format(TBPrograma!Data_fim, "dd/mm/yy")
            ProcGravaEvento
            '==================================
            
            TBPrograma.MoveNext
        Loop
        USMsgBox ("Pedido(s) do produto/serviçoemitido(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Não foi encontrada nenhuma programação com o prazo final menor ou igual a " & Format(DataFim, "dd/mm/yy") & "."), vbExclamation, "CAPRIND v5.0"
    End If
    TBPrograma.Close
End If

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
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If txtCodigo = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtCodigo.SetFocus
    Exit Sub
End If
If txtdescricao = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
'Verifica se o produto já não está na lista
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_programa_item where ID = " & txtId.Text & " and codigo = '" & txtCodigo & "' and ID_item <> " & txtID_item, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Já existe o produto/serviço " & txtCodigo & " nesta programação."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from compras_programa_item where id_item = " & txtID_item & " and id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If txtstatus_item = "PARCIAL" Or txtstatus_item = "RECEBIDO" Then
        USMsgBox ("Não é permitido alterar este produto/serviço, pois o mesmo está " & IIf(txtstatus_item = "RECEBIDO", "recebido", "recebido parcial") & "."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_programacao where id_item = " & TBGravar!Id_Item, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Não é permitido alterar este produto/serviço, pois o mesmo já possui programação."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If
TBGravar!CODIGO = txtCodigo
TBGravar!ID = txtId
If txtstatus_item = "ABERTO" Then TBGravar!Status_Item = "N_RECEBIDO"
TBGravar.Update
txtID_item = TBGravar!Id_Item
TBGravar.Close
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_programa_item where id = " & txtId & " and (status_item = 'RECEBIDO' or status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from compras_programa where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBItem!status = "PARCIAL"
        TBItem.Update
        txtStatus = "PARCIAL"
    End If
    TBItem.Close
End If
TBAbrir.Close
Lista.ListItems.Clear
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
ProcCarregaLista_Item
If Novo_Programacao_Compras1 = True Then
    USMsgBox ("Novo produto/serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto/serviço"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto/serviço"
    If CodigoLista1 <> 0 And lista_item.ListItems.Count <> 0 Then
        lista_item.SelectedItem = lista_item.ListItems(CodigoLista1)
        lista_item.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Programação"
ID_documento = txtID_item
Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
Documento1 = "Cód. interno: " & txtCodigo
ProcGravaEvento
'==================================
Novo_Programacao_Compras1 = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_prog()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame7.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus = "REVISADA" Then
    USMsgBox ("Não é permitido alterar esta programação, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
valor = IIf(txtQTD = "", 0, txtQTD)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    txtQTD.SetFocus
    Exit Sub
End If
With txtData_fim
    If FunVerificaDataFinal(txtData_inicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_programacao where ID_prog = " & txtID_prog, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If txtStatus_prog = "PARCIAL" Or txtStatus_prog = "RECEBIDO" Then
        USMsgBox ("Não é permitido alterar está programação do produto/serviço pois a mesma está " & IIf(txtStatus_prog = "RECEBIDO", "recebido", "recebido parcial") & "."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    ValorTotal = txtQTD - IIf(IsNull(TBGravar!Qtderecebida), 0, TBGravar!Qtderecebida)
    If ValorTotal < 0 Then
        USMsgBox ("Quantidade menor do que já foi recebido."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If TBGravar!Firme = True And optFirme.Value = 1 Then
        USMsgBox ("Não é permitido alterar está programação do produto/serviço pois a mesma já esta firmada."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
End If
ProcEnviadados_Prog
TBGravar.Update
txtID_prog = TBGravar!ID_prog
TBGravar.Close
ProcCarregalista_Prog
If Novo_Programacao_Compras2 = True Then
    USMsgBox ("Nova programação do produto/serviço cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova programação produto/serviço"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar programação produto/serviço"
    If CodigoLista2 <> 0 And lista_prog.ListItems.Count <> 0 Then
        lista_prog.SelectedItem = lista_prog.ListItems(CodigoLista2)
        lista_prog.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Programação"
ID_documento = txtID_prog
Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev & " - Cód. interno: " & txtCodigo & " - Descrição " & txtdescricao
Documento1 = "Data: " & txtData_inicio.Value & "-" & txtData_fim.Value
ProcGravaEvento
'==================================
Novo_Programacao_Compras2 = False

ProcAlteraStatusProgramacao txtID_prog, txtQTD
ProcAltera_Status txtID_item, txtId
optFirme.Enabled = True
Lista.ListItems.Clear
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
ProcCarregaLista_Item

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlteraStatusProgramacao(ID_programacao As Long, Qtde As Double)
On Error GoTo tratar_erro

'Altera status da programação do item
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_programacao where id_prog = " & ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ValorTotal = Qtde - IIf(IsNull(TBAbrir!Qtderecebida), 0, TBAbrir!Qtderecebida)
    If ValorTotal = 0 Then
        TBAbrir!Status_prog = "RECEBIDO"
        TBAbrir!Ordenar = 4
    ElseIf ValorTotal = Qtde Then
            If TBAbrir!Firme = True Then
                TBAbrir!Status_prog = "ABERTO"
                TBAbrir!Ordenar = 2
            Else
                TBAbrir!Status_prog = "PREVISÃO FUTURA"
                TBAbrir!Ordenar = 3
            End If
        Else
            TBAbrir!Status_prog = "PARCIAL"
            TBAbrir!Ordenar = 1
    End If
    TBAbrir.Update
End If
TBAbrir.Close

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
            Case vbKeyF7: ProcRevisar
            Case vbKeyF8: ProcEmitirPC
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_item
            Case vbKeyF3: procSalvar_item
            Case vbKeyF4: procExcluir_item
            Case vbKeyF5: ProcImprimir
            Case vbKeyF8: ProcEmitirPC_item
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_prog
            Case vbKeyF3: ProcSalvar_prog
            Case vbKeyF4: ProcExcluir_prog
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

ProcCarregaToolBar1 Me, 15195, 14, True
ProcCarregaToolBar2 Me, 15195, 11, True
ProcCarregaToolBar3 Me, 15195, 10, True
Formulario = "Compras/Programação"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Programação"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362P" Then
    If USMsgBox("Deseja realmente atualizar os dados das programações?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select Programa, data, Programatexto from Compras_programa order by Programa", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then
            TBCompras.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBCompras.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBCompras.MoveFirst
            Do While TBCompras.EOF = False
                Cont = TBCompras!programa
                Data_Prog = Format(TBCompras!Data, "mm/yyyy")
                a = Cont
                Select Case Len(a)
                    Case 1: a = "000" & Cont & "-" & Data_Prog
                    Case 2: a = "00" & Cont & "-" & Data_Prog
                    Case 3: a = "0" & Cont & "-" & Data_Prog
                    Case 4: a = Cont & "-" & Data_Prog
                End Select
                TBCompras!programatexto = a
                TBCompras.Update
                TBCompras.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from Compras_programacao order by Status_prog", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then
            TBCompras.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBCompras.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBCompras.MoveFirst
            Do While TBCompras.EOF = False
                Select Case TBCompras!Status_prog
                    Case "PARCIAL": TBCompras!Ordenar = 1
                    Case "ABERTO": TBCompras!Ordenar = 2
                    Case "PREVISÃO FUTURA": TBCompras!Ordenar = 3
                    Case "RECEBIDO": TBCompras!Ordenar = 4
                End Select
                TBCompras.Update
                TBCompras.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBCompras.Close
    End If
    USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Compras/Programação"
    Evento = "Atualizar"
    ID_documento = 0
    Documento = ""
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmCompras_programacao_abrir.Show 1

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
                If USMsgBox("Deseja realmente excluir este(s) programa(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Set TBProgramas = CreateObject("adodb.recordset")
            TBProgramas.Open "Select * from Compras_programacao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
            If TBProgramas.EOF = False Then
                Do While TBProgramas.EOF = False
                    If TBProgramas!ID_Antigo <> 0 Then
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from Compras_programacao where id_prog = " & TBProgramas!ID_Antigo, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = False Then
                            TBGravar!Status_prog = "ABERTO"
                            TBGravar!Ordenar = 2
                            TBGravar.Update
                        End If
                        TBGravar.Close
                    End If
                    TBProgramas.MoveNext
                Loop
            End If
            TBProgramas.Close
            
            Conexao.Execute "DELETE from Compras_programa where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from compras_programa_item where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from compras_programacao where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Compras_pedido_lista from Compras_pedido_lista INNER JOIN Compras_pedido ON Compras_pedido_lista.IDpedido = Compras_pedido.IDpedido Where Compras_pedido.ID_programa = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from compras_pedido where ID_programa = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Compras/Programação"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº programação: " & .ListItems(InitFor).SubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
            Set TBPrograma = CreateObject("adodb.recordset")
            TBPrograma.Open "Select * from Compras_programa where Programatexto = '" & .ListItems(InitFor).SubItems(3) & "' and Rev = " & .ListItems(InitFor).SubItems(4) - 1, Conexao, adOpenKeyset, adLockOptimistic
            If TBPrograma.EOF = False Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from compras_programa_item where id = " & TBPrograma!ID & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = True Then
                    TBPrograma!status = "PREVISÃO FUTURA"
                Else
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from compras_programa_item where id = " & TBPrograma!ID & " and Status_Item <> 'N_RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        TBPrograma!status = "ABERTO"
                    Else
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from compras_programa_item where id = " & TBPrograma!ID & " and Status_Item <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = True Then
                            TBPrograma!status = "RECEBIDO"
                        Else
                            TBPrograma!status = "PARCIAL"
                        End If
                    End If
                End If
                TBAbrir.Close
                TBPrograma!Data_rev = Null
                TBPrograma.Update
            End If
            TBPrograma.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) programa(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Programa(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Lista.ListItems.Clear
    ProcCarregaLista (1)
    Framelista.Enabled = False
    ProcLimparTudo
    Novo_Programacao_Compras = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtPrograma = "" Then
    USMsgBox ("Informe o programa antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Dataini = 0
DataFim = 0
'Verifica a datainicio e final das programações q estão e não estão como recebidas
Set TBProgramas = CreateObject("adodb.recordset")
If txtStatus <> "ENCERRADO" Then TextoFiltro = "Status_prog <> 'RECEBIDO'" Else TextoFiltro = "Status_prog = 'RECEBIDO'"
TBProgramas.Open "Select * from Compras_Programacao where ID = " & txtId & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProgramas.EOF = False Then
    Dataini = TBProgramas!Data_inicio
    TBProgramas.MoveLast
    DataFim = TBProgramas!Data_fim
End If
TBProgramas.Close
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_Programa where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!Dtinicio_rel = Dataini
    TBGravar!Dtfinal_rel = DataFim
    'TBGravar!via = IIf(IsNull(TBGravar!via), 0, TBGravar!via) + 1
    TBGravar.Update
End If
TBGravar.Close

NomeRel = "Compras_programacao.rpt"
ProcImprimirRel "{Compras_Programa.id} = " & txtId, ""

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
Data_Prog = ""
ProcLimpaCampos
Novo_Programacao_Compras = True
Framelista.Enabled = True
txtData = Format(Date, "dd/mm/yy")
txtStatus = "ABERTO"
cmdpesquisar_Click
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame4.Enabled = False
Frame7.Enabled = False
ProcLimpaCampos_Item
ProcLimpaCampos_Prog
Novo_Programacao_Compras1 = False
Novo_Programacao_Compras2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Programacao_Compras = True Then
    If USMsgBox("O programa ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Programacao_Compras = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Programacao_Compras1 = True Then
    If USMsgBox("O produto/serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_item
        If Novo_Programacao_Compras1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Programacao_Compras2 = True Then
    If USMsgBox("A programação do produto/serviço ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_prog
        If Novo_Programacao_Compras2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Programacao_Compras = False
Novo_Programacao_Compras1 = False
Novo_Programacao_Compras2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtPrograma = ""
txtFornecedor = ""
txtID_forn = ""
txtResponsavel = pubUsuario
txtStatus = "ABERTO"
txtData.Text = Format(Date, "dd/mm/yy")
txtData_rev.Text = ""
txtRev = "0"
CodigoLista = 0
Caption = "Administrativo - Compras - Programação"

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
If Framelista.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus = "RECEBIDO" Or txtStatus = "PARCIAL" Or txtStatus = "REVISADA" Then
    Select Case txtStatus
        Case "RECEBIDO": Mensagem = "recebido"
        Case "PARCIAL": Mensagem = "recebido parcial"
        Case "REVISADA": Mensagem = "revisado"
    End Select
    USMsgBox ("Não é permitido alterar este programa, pois o mesmo está " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "Salvar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtID_forn = "" Or txtFornecedor = "" Then
    NomeCampo = "o fornecedor"
    ProcVerificaAcao
    Exit Sub
End If
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * FROM Compras_fornecedores WHERE idcliente = " & txtID_forn, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    With Cmb_empresa
        If FunVerifValidadeCertForn(.ItemData(.ListIndex), txtData, True) = False Then Exit Sub
        If TBFornecedor!Pessoa = "JURÍDICA" And TBFornecedor!idTipoEmpresa = 1 Then
            If FunVerifRegimeTribCliForn(.ItemData(.ListIndex), True, True) = False Then Exit Sub
        End If
    End With
End If
TBFornecedor.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Compras_programa where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_programa order by id", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = True Then
        Cont = 1
    Else
        TBCompras.MoveLast
        Cont = TBCompras!programa + 1
    End If
    Data_Prog = Format(Date, "mm/yyyy")
    ProcGeraNumero
    TBCompras.Close

    TBGravar.AddNew
    TBGravar!programa = Cont
    TBGravar!programatexto = a
    TBGravar!Data = Date
    TBGravar!Responsavel = txtResponsavel
    TBGravar!status = "ABERTO"
End If
TBGravar!ID_forn = txtID_forn
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar.Update
txtId = TBGravar!ID
txtPrograma.Text = a
TBGravar.Close

Lista.ListItems.Clear
If Novo_Programacao_Compras = True Then
    USMsgBox ("Novo programa cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Programacao_Compras_Localizar = "Select compras_programa.*, Compras_fornecedores.Nome_Razao FROM compras_programa INNER JOIN Compras_fornecedores ON compras_programa.id_Forn = Compras_fornecedores.IDCliente where compras_programa.ID = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Programação"
ID_documento = txtId
Documento = "Nº programa: " & txtPrograma & " - Rev.: " & txtRev
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Programacao_Compras = False

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
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Compras_programa where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!status = "ENCERRADO" Or TBAbrir!status = "PARCIAL" Or TBAbrir!status = "REVISADA" Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                End If
                TBAbrir.Close
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

Private Sub lista_item_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_item
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtStatus = "REVISADA" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If .ListItems(InitFor).SubItems(3) = "PARCIAL" Or .ListItems(InitFor).SubItems(3) = "RECEBIDO" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_item, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_item_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_item
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If txtStatus = "REVISADA" Then
                USMsgBox ("Não é permitido excluir este produto/serviço, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
'=================================================================================================================
' Verifica status do item
'=================================================================================================================
Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Compras_programa_item where id_item = " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!Status_Item = "RECEBIDO" Or TBAbrir!Status_Item = "PARCIAL" Then
                    Select Case TBAbrir!Status_Item
                        Case "RECEBIDO": Mensagem = "recebido"
                        Case "PARCIAL": Mensagem = "comprado parcial"
                    End Select
                    USMsgBox ("Não é permitido excluir este produto/serviço, pois o mesmo está " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
            End If
            TBAbrir.Close
'=================================================================================================================
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_item_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_item.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos_Item
txtID_item = lista_item.SelectedItem
txtCodigo = lista_item.SelectedItem.ListSubItems(1)
txtdescricao = lista_item.SelectedItem.ListSubItems(2)
txtstatus_item = lista_item.SelectedItem.ListSubItems(3)
CodigoLista1 = lista_item.SelectedItem.index
Frame4.Enabled = True
Novo_Programacao_Compras1 = False

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
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Compras_programa where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!status = "ENCERRADO" Or TBAbrir!status = "PARCIAL" Or TBAbrir!status = "REVISADA" Then
                    Select Case TBAbrir!status
                        Case "ENCERRADO": Mensagem = "recebido"
                        Case "PARCIAL": Mensagem = "comprado parcial"
                        Case "REVISADA": Mensagem = "revisado"
                    End Select
                    USMsgBox ("Não é permitido excluir este programa, pois o mesmo está " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_programa where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBAbrir.Close
Framelista.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then
    ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
    Empresarel = Cmb_empresa
End If
Cont = TBAbrir!programa
Caption = "Administrativo - Compras - Programação - (Programação : " & IIf(IsNull(TBAbrir!programatexto), "", TBAbrir!programatexto) & " - Rev. : " & IIf(IsNull(TBAbrir!Rev), "", TBAbrir!Rev) & ")"
txtId = TBAbrir!ID
txtPrograma = TBAbrir!programatexto
txtID_forn = TBAbrir!ID_forn
txtResponsavel = TBAbrir!Responsavel
txtStatus = TBAbrir!status
txtData = Format(TBAbrir!Data, "dd/mm/yy")
txtRev = IIf(IsNull(TBAbrir!Rev), "0", TBAbrir!Rev)
txtData_rev = IIf(IsNull(TBAbrir!Data_rev), "", Format(TBAbrir!Data_rev, "dd/mm/yy"))
Novo_Programacao_Compras = False
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_nota_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lista_nota, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_prog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_prog
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtStatus = "REVISADA" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
            
                If .ListItems(InitFor).SubItems(6) <> "ABERTO" And .ListItems(InitFor).SubItems(6) <> "PREVISÃO FUTURA" Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_prog, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_prog_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_prog
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If txtStatus = "REVISADA" Then
                USMsgBox ("Não é permitido excluir esta programação, pois o programa está revisado."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            
            If .ListItems(InitFor).SubItems(6) <> "ABERTO" And .ListItems(InitFor).SubItems(6) <> "PREVISÃO FUTURA" Then
                USMsgBox ("Não é permitido excluir esta programação do produto/serviço pois a mesma está " & IIf(.ListItems(InitFor).SubItems(6) = "RECEBIDO", "recebida", "recebida parcial") & "."), vbExclamation, "CAPRIND v5.0"
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

Private Sub lista_prog_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_prog.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_programacao where ID_prog = " & lista_prog.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos_Prog
    
    txtID_prog = TBAbrir!ID_prog
    txtData_inicio.Value = TBAbrir!Data_inicio
    txtData_fim.Value = TBAbrir!Data_fim
    txtQTD = Format(TBAbrir!quantidade, "###,##0.0000")
    txtStatus_prog.Text = TBAbrir!Status_prog
    If TBAbrir!Firme = True Then optFirme.Value = 1 Else optFirme.Value = 0
    Txt_un = IIf(IsNull(TBAbrir!Un), "", TBAbrir!Un)
    Txt_un_com = IIf(IsNull(TBAbrir!Unidade_com), "", TBAbrir!Unidade_com)
    
    Frame7.Enabled = True
    CodigoLista2 = lista_prog.SelectedItem.index
    Novo_Programacao_Compras2 = False
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFirme_Click()
On Error GoTo tratar_erro

With lista_prog
'=================================================================================================================
' Verifica status do item na programação
'=================================================================================================================
Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Compras_programacao where id_prog = " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!Status_prog = "RECEBIDO" Or TBAbrir!Status_prog = "PARCIAL" Then
                    Select Case TBAbrir!Status_prog
                        Case "RECEBIDO": Mensagem = "recebido"
                        Case "PARCIAL": Mensagem = "comprado parcial"
                    End Select
                    USMsgBox ("Não é permitido cancelar essa programação desse produto/serviço, pois o mesmo está " & Mensagem & "."), vbExclamation, "CAPRIND v5.0"
                    ProcCarregalista_Prog
                    optFirme.Value = 1
                End If
            End If
            TBAbrir.Close
'=================================================================================================================
       ' End If
'    Next InitFor
End With

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
        Cmb_empresa.Visible = True
        PBLista.Visible = True
        If Cmb_empresa.Visible = True Then Cmb_empresa.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        PBLista.Visible = True
        ProcCarregaFornecedor
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        lista_item.SetFocus
        ProcCarregaLista_Item
    Case 2:
        Cmb_empresa.Visible = False
        PBLista.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        If txtCodigo = "" Then
            SSTab1.Tab = 1
            USMsgBox ("Informe o produto/serviço antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        If Novo_Programacao_Compras1 = True Then
            SSTab1.Tab = 1
            USMsgBox ("Salve o produto/serviço antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            Permitido = False
            Exit Sub
        End If
        lista_prog.SetFocus
        txtCodigo2 = txtCodigo
        txtDescricao2 = txtdescricao
        ProcVerificaUnidade
        ProcCarregalista_Prog
        txtTotal.Text = "0,0000"
        Proccarreganota
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Programacao_Compras = True Then
    SSTab1.Tab = 0
    USMsgBox ("Salve o programa antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    Exit Sub
End If
If txtFornecedor = "" Then
    SSTab1.Tab = 0
    USMsgBox ("Informe o fornecedor antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    cmdpesquisar_Click
    Permitido = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaUnidade()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Unidade, Unidade_com from projproduto where desenho = '" & txtCodigo2.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Txt_un.Text = TBProduto!Unidade
    Txt_un_com.Text = TBProduto!Unidade_com
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

If Sql_Programacao_Compras_Localizar = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_Compras_Programacao = CreateObject("adodb.recordset")
TBLISTA_Compras_Programacao.Open Sql_Programacao_Compras_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Compras_Programacao.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Compras_Programacao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Compras_Programacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Compras_Programacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Compras_Programacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Compras_Programacao.PageSize * (Pagina - 1)), 0), TBLISTA_Compras_Programacao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Compras_Programacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Compras_Programacao!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Compras_Programacao!Data), "", Format(TBLISTA_Compras_Programacao!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Compras_Programacao!Responsavel), "", TBLISTA_Compras_Programacao!Responsavel)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Compras_Programacao!programatexto), "", TBLISTA_Compras_Programacao!programatexto)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Compras_Programacao!Rev), 0, TBLISTA_Compras_Programacao!Rev)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Compras_Programacao!Data_rev), "", Format(TBLISTA_Compras_Programacao!Data_rev, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Compras_Programacao!Nome_Razao), "", TBLISTA_Compras_Programacao!Nome_Razao)
        .Item(.Count).SubItems(7) = IIf(TBLISTA_Compras_Programacao!status = "ENCERRADO", "RECEBIDO", TBLISTA_Compras_Programacao!status)
    End With
    TBLISTA_Compras_Programacao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Compras_Programacao.RecordCount
If TBLISTA_Compras_Programacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Compras_Programacao.PageCount
ElseIf TBLISTA_Compras_Programacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Compras_Programacao.PageCount & " de: " & TBLISTA_Compras_Programacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Compras_Programacao.AbsolutePage - 1 & " de: " & TBLISTA_Compras_Programacao.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Item()
On Error GoTo tratar_erro

txtID_item.Text = 0
txtCodigo.Text = ""
txtstatus_item = "ABERTO"
txtdescricao.Text = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Item()
On Error GoTo tratar_erro

lista_item.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from compras_programa_item where id = " & txtId & " order by codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    With lista_item.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!Id_Item
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select * from projproduto where desenho = '" & TBLISTA!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
            End If
            TBItem.Close
            If TBLISTA!Status_Item = "N_RECEBIDO" Then status = "ABERTO" Else status = TBLISTA!Status_Item
            .Item(.Count).SubItems(3) = status
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End With
End If
TBLISTA.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGeraNumero()
On Error GoTo tratar_erro

a = Cont
Select Case Len(a)
    Case 1: a = "000" & Cont & "-" & Data_Prog
    Case 2: a = "00" & Cont & "-" & Data_Prog
    Case 3: a = "0" & Cont & "-" & Data_Prog
    Case 4: a = Cont & "-" & Data_Prog
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Prog()
On Error GoTo tratar_erro

txtID_prog = 0
txtData_fim.Value = Date
txtData_inicio.Value = Date
txtStatus_prog = "ABERTO"
txtQTD = ""
ProcVerificaUnidade
optFirme.Value = 0
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

txtdescricao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtID_forn_Change()
On Error GoTo tratar_erro

If txtID_forn <> "" Then
    VerifNumero = txtID_forn
    ProcVerificaNumero
    If VerifNumero = False Then
        txtID_forn = ""
        txtID_forn.SetFocus
        Exit Sub
    End If
End If
ProcCarregaFornecedor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtID_forn_Click()
On Error GoTo tratar_erro

If txtID_forn = "0" Then txtID_forn = ""

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

Private Sub txtQtd_Change()
On Error GoTo tratar_erro

If txtQTD.Text <> "" Then
    VerifNumero = txtQTD.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQTD.Text = ""
        txtQTD.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtd_LostFocus()
On Error GoTo tratar_erro

txtQTD.Text = Format(txtQTD.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadados_Prog()
On Error GoTo tratar_erro

TBGravar!ID = txtId
TBGravar!Id_Item = txtID_item
TBGravar!Un = Txt_un
TBGravar!Unidade_com = Txt_un_com
TBGravar!Data_inicio = Format(txtData_inicio.Value, "dd/mm/yyyy")
TBGravar!Data_fim = Format(txtData_fim.Value, "dd/mm/yyyy")
TBGravar!quantidade = Format(txtQTD, "###,##0.0000")
If txtStatus_prog <> "PARCIAL" And txtStatus_prog <> "RECEBIDO" Then
    If optFirme.Value = 1 Then
        TBGravar!Firme = True
        TBGravar!Status_prog = "ABERTO"
        TBGravar!Ordenar = 2
        txtStatus_prog = "ABERTO"
    Else
        TBGravar!Firme = False
        TBGravar!Status_prog = "PREVISÃO FUTURA"
        TBGravar!Ordenar = 3
        txtStatus_prog = "PREVISÃO FUTURA"
    End If
    
    If TBGravar!Firme = True Then
        ProcGerarPedido txtID_prog, txtData_fim, txtCodigo, txtdescricao, txtQTD
    Else
        If txtID_prog <> 0 Then ProcVerifExcluirPedido "Compras_pedido_lista.Id_programacao = " & txtID_prog
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarPedido(ID_programacao As Long, Prazo_final As Date, Codinterno As String, Descricao As String, Qtde As Double)
On Error GoTo tratar_erro

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from Compras_pedido where ID_programa = " & txtId & " and Data = '" & Date & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = True Then
    TBPedido.AddNew
    TBPedido!Data = Date
    TBPedido!Responsavel = pubUsuario
    TBPedido!Status_pedido = "AGUARDANDO APROVAÇÃO"
    
    'Gerar numero do pedido de compra
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido where Year(data) = '" & Year(Date) & "' order by IDPedido", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir.MoveLast
        Numero = Left(TBAbrir!Pedido, Len(TBAbrir!Pedido) - 3) + 1
    Else
        Numero = 1
    End If
    TBAbrir.Close
    Ano = Right(Year(Date), 2)
    TBPedido!Pedido = Numero & "/" & Ano
End If
            
TBPedido!ID_programa = txtId
TBPedido!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBPedido!IDFornecedor = txtID_forn
                            
'Gravar dados do fornecedor
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & txtID_forn, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    TBPedido!Fornecedor = TBFornecedor!Nome_Razao
    TBPedido!Categoria = IIf(IsNull(TBFornecedor!Categoria), "A", TBFornecedor!Categoria)
    TBPedido!Tipo_endereco = IIf(IsNull(TBFornecedor!Tipo_endereco), Null, TBFornecedor!Tipo_endereco)
    TBPedido!Endereco = IIf(IsNull(TBFornecedor!Endereco), Null, TBFornecedor!Endereco)
    TBPedido!Numero = IIf(IsNull(TBFornecedor!Numero), Null, TBFornecedor!Numero)
    TBPedido!Tipo_bairro = IIf(IsNull(TBFornecedor!Tipo_bairro), Null, TBFornecedor!Tipo_bairro)
    TBPedido!Bairro = IIf(IsNull(TBFornecedor!Bairro), Null, TBFornecedor!Bairro)
    TBPedido!Cidade = IIf(IsNull(TBFornecedor!Cidade), Null, TBFornecedor!Cidade)
    TBPedido!Estado = IIf(IsNull(TBFornecedor!Estado), Null, TBFornecedor!Estado)
    TBPedido!Email = IIf(IsNull(TBFornecedor!Email), Null, TBFornecedor!Email)
    TBPedido!fone = IIf(IsNull(TBFornecedor!Telefones), Null, TBFornecedor!Telefones)
    TBPedido!Fax = IIf(IsNull(TBFornecedor!Fax), Null, TBFornecedor!Fax)
End If
TBFornecedor.Close

TBPedido!Status_pedido = "AGUARDANDO APROVAÇÃO"
TBPedido.Update
IDpedido = TBPedido!IDpedido
TBPedido.Close

'Gravar produtos
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from Compras_pedido_lista where IDpedido = " & IDpedido & " and ID_programacao = " & ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = True Then
    TBCompras_Lista.AddNew
    TBCompras_Lista!ID_programacao = ID_programacao
    TBCompras_Lista!Remessa = False
    TBCompras_Lista!Status_Item = "AGUARDANDO APROVAÇÃO"
End If
TBCompras_Lista!IDpedido = IDpedido

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBCompras_Lista!Codproduto = TBProduto!Codproduto
    
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Codproduto = " & TBProduto!Codproduto & " and IA.ID_cliente_forn = " & txtID_forn & " and IA.Tipo = 'C' and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = True Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Codproduto = " & TBProduto!Codproduto & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBFIltro.EOF = False Then
        TBCompras_Lista!N_referencia = TBFIltro!N_referencia
    End If
    TBFIltro.Close
    
    TBCompras_Lista!Descricao_comercial = TBProduto!descricaotecnica
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Projproduto_fornecedor where Codproduto = " & TBProduto!Codproduto & " and idfornecedor = " & txtID_forn, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        TBCompras_Lista!preco_unitario = IIf(IsNull(TBFI!PCusto), "", Format(TBFI!PCusto, "###,##0.0000000000"))
    Else
        TBCompras_Lista!preco_unitario = IIf(IsNull(TBProduto!PCusto), "", (Format(TBProduto!PCusto, "###,##0.0000000000")))
    End If
    TBFI.Close
    
    TBCompras_Lista!preco_unitario_desconto = Format(TBCompras_Lista!preco_unitario, "###,##0.0000000000")
    
    ProcAgregarProdutoForn TBProduto!Codproduto, txtID_forn, TBCompras_Lista!preco_unitario
    
    TBCompras_Lista!Familia = TBProduto!Classe
    TBCompras_Lista!Un = TBProduto!Unidade
    TBCompras_Lista!Unidade_com = TBProduto!Unidade_com
End If
TBCompras_Lista!Desenho = Codinterno
TBCompras_Lista!Descricao = Descricao
TBCompras_Lista!Quant_Comp = Qtde
TBCompras_Lista!Quant_Comp_PC = FunCalculaQtdePC(Codinterno, Qtde, True, TBCompras_Lista!Unidade_com)
TBCompras_Lista!preco_total = Format(TBCompras_Lista!preco_unitario * TBCompras_Lista!Quant_Comp, "###,##0.00")
TBCompras_Lista!Prazo = Prazo_final
TBCompras_Lista!Tipo = "P"
TBCompras_Lista.Update

Totalpedido = 0
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select idpedido, dbl_valor_total_produtos, dbl_valor_total_servicos, TotalDesconto, SubTotal, dbl_Valor_Total_IPI, dbl_valor_total from compras_pedido where IDpedido = " & IDpedido & " order by idpedido", Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select Sum(preco_total) as Totalpedido from compras_pedido_lista where idpedido = " & TBPedido!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        Totalpedido = IIf(IsNull(TBCompras_Lista!Totalpedido), 0, TBCompras_Lista!Totalpedido)
    End If
    TBCompras_Lista.Close
    TBPedido!dbl_Valor_Total_Produtos = Totalpedido
    TBPedido!SubTotal = (Totalpedido + TBPedido!dbl_valor_total_servicos) - TBPedido!TotalDesconto
    TBPedido!dbl_valor_total = TBPedido!SubTotal + TBPedido!dbl_Valor_Total_IPI
    TBPedido.Update
End If
TBPedido.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifExcluirPedido(TextoFiltro As String)
On Error GoTo tratar_erro

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select compras_pedido.Idpedido, compras_pedido.dbl_valor_total, compras_pedido_lista.IDlista from compras_pedido_lista INNER JOIN compras_pedido ON compras_pedido_lista.IDpedido = compras_pedido.Idpedido where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
    Conexao.Execute "DELETE from Compras_pedido_lista from Compras_pedido_lista INNER JOIN Compras_pedido ON Compras_pedido_lista.IDpedido = Compras_pedido.IDpedido Where " & TextoFiltro
    Conexao.Execute "DELETE from Compras_pedido_lista_custo WHERE IDLista = " & TBPedido!IDlista
    
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select * from compras_pedido_lista where IDpedido = " & TBPedido!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = True Then
        Conexao.Execute "DELETE compras_pedido where idpedido = " & TBPedido!IDpedido
        Conexao.Execute "DELETE Compras_comercial where idpedido = " & TBPedido!IDpedido
    Else
        Totalpedido = 0
        Set TBCompras_Lista = CreateObject("adodb.recordset")
        TBCompras_Lista.Open "Select Sum(preco_total) as Totalpedido from compras_pedido_lista where idpedido = " & TBPedido!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras_Lista.EOF = False Then
            Totalpedido = IIf(IsNull(TBCompras_Lista!Totalpedido), 0, TBCompras_Lista!Totalpedido)
        End If
        TBCompras_Lista.Close
        TBPedido!dbl_valor_total = Totalpedido
        TBPedido.Update
    End If
End If
TBPedido.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Prog()
On Error GoTo tratar_erro

lista_prog.ListItems.Clear
ValorTotal = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_programacao where id = " & txtId & " and Id_item = " & txtID_item & " order by Ordenar, data_inicio desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista2.Min = 0
    PBLista2.Max = TBLISTA.RecordCount
    PBLista2.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    With lista_prog.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID_prog
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data_inicio), "", Format(TBLISTA!Data_inicio, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data_fim), "", Format(TBLISTA!Data_fim, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!quantidade), "0,0000", Format(TBLISTA!quantidade, "###,##0.0000"))
            ValorTotal = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade) - IIf(IsNull(TBLISTA!Qtderecebida), 0, TBLISTA!Qtderecebida)
            .Item(.Count).SubItems(4) = Format(ValorTotal, "###,##0.0000")
            .Item(.Count).SubItems(5) = IIf(TBLISTA!Firme = True, "Sim", "Não")
            If TBLISTA!Status_prog = "ENCERRADO" Then status = "RECEBIDO" Else status = IIf(IsNull(TBLISTA!Status_prog), "", TBLISTA!Status_prog)
            .Item(.Count).SubItems(6) = status
            
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select compras_pedido.Pedido from compras_pedido INNER JOIN compras_pedido_lista ON compras_pedido.IDpedido = compras_pedido_lista.IDpedido where compras_pedido_lista.ID_programacao = " & TBLISTA!ID_prog, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                .Item(.Count).SubItems(7) = IIf(IsNull(TBPedido!Pedido), "", TBPedido!Pedido)
            End If
            TBPedido.Close
            
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista2.Value = Contador
        Loop
    End With
End If
TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub Proccarreganota()
On Error GoTo tratar_erro

Quant = 0
lista_nota.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Estoque_controle_recebimento where idlista = " & txtID_item & " and idpedido = " & txtId & " and Programacao = 'True' order by Nota_fiscal", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcCarregaListaNF
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Estoque_controle_recebimento.* from (((Estoque_controle_recebimento INNER JOIN Compras_pedido_lista ON Estoque_controle_recebimento.IDlista = Compras_pedido_lista.IDlista and Estoque_controle_recebimento.Idpedido = Compras_pedido_lista.IDpedido) INNER JOIN Compras_Programacao ON Compras_Programacao.ID_prog = Compras_pedido_lista.ID_programacao) INNER JOIN Compras_programa_item ON Compras_programa_item.ID_item = Compras_Programacao.ID_item) INNER JOIN Compras_programa ON Compras_programa.ID = Compras_programa_item.ID where Compras_programa_item.codigo = '" & txtCodigo & "' and Compras_programa.Programatexto = '" & txtPrograma & "' and Estoque_controle_recebimento.Programacao = 'False' order by Estoque_controle_recebimento.Nota_fiscal", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcCarregaListaNF
End If
TBLISTA.Close
txtTotal = Format(Quant, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNF()
On Error GoTo tratar_erro

TBLISTA.MoveLast
PBLista3.Min = 0
PBLista3.Max = TBLISTA.RecordCount
PBLista3.Value = 1
Contador = 0
TBLISTA.MoveFirst
Do While TBLISTA.EOF = False
    With lista_nota.ListItems
        .Add , , TBLISTA!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Nota_fiscal), "", TBLISTA!Nota_fiscal)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data_emissao), "", Format(TBLISTA!Data_emissao, "dd/mm/yy"))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Recebido), "0,0000", Format(TBLISTA!Recebido, "###,##0.0000"))
        Quant = Quant + IIf(IsNull(TBLISTA!Recebido), "0,0000", Format(TBLISTA!Recebido, "###,##0.0000"))
    End With
    TBLISTA.MoveNext
    Contador = Contador + 1
    PBLista3.Value = Contador
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAltera_Status(ID_prod As Long, ID_programa As Long)
On Error GoTo tratar_erro

'Prooduto
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from compras_programa_item where id_item = " & ID_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_programacao where id_item = " & ID_prod & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!Status_Item = "PREVISÃO FUTURA"
        txtstatus_item = "PREVISÃO FUTURA"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from compras_programacao where id_item = " & ID_prod & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!Status_Item = "N_RECEBIDO"
            txtstatus_item = "ABERTO"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_programacao where id_item = " & ID_prod & " and status_prog <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!Status_Item = "RECEBIDO"
                txtstatus_item = "RECEBIDO"
            Else
                TBItem!Status_Item = "PARCIAL"
                txtstatus_item = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    TBItem.Update
End If

'Programa
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from compras_programa where id = " & ID_programa, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_programa_item where id = " & ID_programa & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBItem!status = "PREVISÃO FUTURA"
        txtStatus = "PREVISÃO FUTURA"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from compras_programa_item where id = " & ID_programa & " and Status_Item <> 'N_RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            TBItem!status = "ABERTO"
            txtStatus = "ABERTO"
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_programa_item where id = " & ID_programa & " and Status_Item <> 'RECEBIDO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                TBItem!status = "RECEBIDO"
                txtStatus = "RECEBIDO"
            Else
                TBItem!status = "PARCIAL"
                txtStatus = "PARCIAL"
            End If
        End If
    End If
    TBAbrir.Close
    TBItem.Update
End If
TBItem.Close
ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
ProcCarregaLista_Item

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaFornecedor()
On Error GoTo tratar_erro

txtFornecedor = ""
If txtID_forn = "" Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
If Novo_Programacao_Compras = True Then TextoFiltro = "idcliente = " & txtID_forn & " and status <> 'Bloqueado' and Prospecto = 'False'" Else TextoFiltro = "idcliente = " & txtID_forn
TBFornecedor.Open "Select * from compras_fornecedores where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    txtFornecedor = IIf(IsNull(TBFornecedor!Nome_Razao), "", TBFornecedor!Nome_Razao)
End If
TBFornecedor.Close

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
    Case 8: ProcRevisar
    Case 9: procAtualiza
    Case 10: ProcEmitirPC
    Case 12: ProcAjuda
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
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcEmitirPC_item
    Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_prog
    Case 2: ProcSalvar_prog
    Case 3: ProcExcluir_prog
    Case 4: ProcImprimir
    Case 5: ProcAnterior_prog
    Case 6: ProcProximo_prog
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
