VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_ListaProduto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Vendas | Proposta comercial | Localizar"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   17715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendas_ListaProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   17715
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   40
      Top             =   9345
      Width           =   17715
      _ExtentX        =   31247
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   17715
      _ExtentX        =   31247
      _ExtentY        =   741
      DibPicture      =   "frmVendas_ListaProduto.frx":0CCA
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_ListaProduto.frx":431A
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
      FormHeightDT    =   9750
      FormWidthDT     =   17715
      FormScaleHeightDT=   9750
      FormScaleWidthDT=   17715
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   135
      Left            =   30
      TabIndex        =   23
      Top             =   9900
      Width           =   17475
      _ExtentX        =   30824
      _ExtentY        =   238
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
      Height          =   8880
      Left            =   60
      TabIndex        =   16
      Top             =   480
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   15663
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
      TabCaption(0)   =   "Lista de produtos"
      TabPicture(0)   =   "frmVendas_ListaProduto.frx":4FF4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "btnAdicionar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GridItens"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnFiltrar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Estrutura"
      TabPicture(1)   =   "frmVendas_ListaProduto.frx":5010
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "USImageList2"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "Grid1"
      Tab(1).ControlCount=   4
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   855
         Left            =   14610
         TabIndex        =   43
         ToolTipText     =   "Buscar itens"
         Top             =   330
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1508
         DibPicture      =   "frmVendas_ListaProduto.frx":502C
         Caption         =   "Buscar (F2)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin FlexCell.Grid GridItens 
         Height          =   6915
         Left            =   60
         TabIndex        =   32
         Top             =   1200
         Width           =   17445
         _ExtentX        =   30771
         _ExtentY        =   12197
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor2      =   14737632
         BackColorActiveCellSel=   12640511
         BackColorBkg    =   16777215
         BorderColor     =   12632256
         CellBorderColor =   8421504
         SelectionBorderColor=   4210752
         Cols            =   12
         DefaultFontSize =   8.25
         DisplayFocusRect=   0   'False
         DisplayRowIndex =   -1  'True
         FixedRowColStyle=   2
         GridColor       =   12632256
         ReadOnlyFocusRect=   0
         Rows            =   1
         ScrollBars      =   2
         ScrollBarStyle  =   0
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
         AllowUserPaste  =   2
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -74940
         TabIndex        =   28
         Top             =   1290
         Width           =   17445
         Begin VB.ComboBox cmbVersao_pesquisar_estrutura 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmVendas_ListaProduto.frx":867C
            Left            =   2070
            List            =   "frmVendas_ListaProduto.frx":86CE
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Versão."
            Top             =   210
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisa por versão :"
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
            Index           =   19
            Left            =   180
            TabIndex        =   29
            Top             =   210
            Width           =   1800
         End
         Begin VB.Image imgFolder 
            Height          =   240
            Left            =   10560
            Picture         =   "frmVendas_ListaProduto.frx":8720
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgFile 
            Height          =   240
            Left            =   10830
            Picture         =   "frmVendas_ListaProduto.frx":8CAA
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   615
         Left            =   55
         TabIndex        =   20
         Top             =   8130
         Width           =   17445
         Begin VB.ComboBox Cmb_ordenar 
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
            ItemData        =   "frmVendas_ListaProduto.frx":9234
            Left            =   16110
            List            =   "frmVendas_ListaProduto.frx":923E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            ToolTipText     =   "Ordenar por."
            Top             =   210
            Width           =   1185
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
            Left            =   3930
            TabIndex        =   7
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
            Left            =   8220
            TabIndex        =   8
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   10440
            TabIndex        =   12
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_ListaProduto.frx":925D
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
            Left            =   9900
            TabIndex        =   11
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_ListaProduto.frx":CA01
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
            Left            =   8790
            TabIndex        =   9
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
            Left            =   9360
            TabIndex        =   10
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_ListaProduto.frx":1050A
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
            Left            =   10980
            TabIndex        =   13
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_ListaProduto.frx":145F9
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenar por"
            BeginProperty Font 
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
            Left            =   15150
            TabIndex        =   42
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label2 
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
            Left            =   4560
            TabIndex        =   30
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label24 
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
            Left            =   3240
            TabIndex        =   25
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
            Left            =   12090
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   55
         TabIndex        =   17
         Top             =   330
         Width           =   14535
         Begin VB.TextBox Txt_dias 
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
            Left            =   13080
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Prazo final."
            Top             =   390
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CheckBox Chk_prazo_todos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prazo (Todos)"
            BeginProperty Font 
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
            Left            =   13080
            TabIndex        =   37
            Top             =   180
            Width           =   1365
         End
         Begin VB.TextBox Txt_pedido_cliente 
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
            Left            =   11730
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Pedido do cliente."
            Top             =   390
            Width           =   1335
         End
         Begin VB.CheckBox Chk_pedido_todos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pedido todos"
            BeginProperty Font 
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
            Left            =   11760
            TabIndex        =   35
            Top             =   180
            Width           =   2115
         End
         Begin VB.CheckBox Chk_carregarinfest 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Carregar estoque"
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
            Left            =   9570
            TabIndex        =   34
            Top             =   270
            Width           =   1605
         End
         Begin VB.CheckBox Chk_cliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pertencente(s) ao cliente"
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
            Left            =   9570
            TabIndex        =   33
            Top             =   510
            Width           =   2505
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   2550
            TabIndex        =   27
            Top             =   240
            Width           =   2895
            Begin VB.OptionButton Optfim 
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
               Height          =   255
               Left            =   1560
               TabIndex        =   5
               Top             =   180
               Width           =   555
            End
            Begin VB.OptionButton Optinicio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   3
               Top             =   180
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton Optmeio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   4
               Top             =   180
               Width           =   645
            End
            Begin VB.OptionButton optIgual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2130
               TabIndex        =   6
               Top             =   180
               Width           =   705
            End
         End
         Begin VB.TextBox txtTexto 
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
            Left            =   5490
            TabIndex        =   1
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   3975
         End
         Begin VB.ComboBox cmbfiltrarpor 
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
            ItemData        =   "frmVendas_ListaProduto.frx":17E85
            Left            =   180
            List            =   "frmVendas_ListaProduto.frx":17EB0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   2325
         End
         Begin VB.ComboBox cmbfamilia 
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
            Left            =   5520
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Visible         =   0   'False
            Width           =   3945
         End
         Begin MSComCtl2.DTPicker Cmb_prazo 
            Height          =   315
            Left            =   13170
            TabIndex        =   39
            ToolTipText     =   "Prazo final."
            Top             =   390
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   180617219
            CurrentDate     =   39057
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filtrar por"
            BeginProperty Font 
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
            Left            =   915
            TabIndex        =   19
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Texto para pesquisa"
            BeginProperty Font 
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
            Left            =   6735
            TabIndex        =   18
            Top             =   180
            Width           =   1470
         End
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -70950
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_ListaProduto.frx":17F67
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74940
         TabIndex        =   24
         Top             =   330
         Width           =   17445
         _ExtentX        =   30771
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
         ButtonKey3      =   "3"
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7035
         Left            =   -74940
         TabIndex        =   15
         Top             =   1980
         Width           =   17445
         _ExtentX        =   30771
         _ExtentY        =   12409
         Cols            =   2
         DefaultFontSize =   8.25
         GridColor       =   12632256
         ReadOnly        =   -1  'True
         Rows            =   2
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6045
         Left            =   60
         TabIndex        =   26
         Top             =   2070
         Width           =   17445
         _ExtentX        =   30771
         _ExtentY        =   10663
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   8683
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição comercial"
            Object.Width           =   8683
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Emp. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Disp. est."
            Object.Width           =   1587
         EndProperty
      End
      Begin DrawSuite2022.USButton btnAdicionar 
         Height          =   855
         Left            =   16080
         TabIndex        =   44
         ToolTipText     =   "Adicionar itens selecionados"
         Top             =   330
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1508
         DibPicture      =   "frmVendas_ListaProduto.frx":1958C
         Caption         =   "Adicionar (F3)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
End
Attribute VB_Name = "frmVendas_ListaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSqlLocProdVendas As String 'OK
Dim FlexGrid_Desenho    As String 'ok
Dim PermitidoColuna     As Boolean 'OK

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, DataValidacao As String, RespValidacao As String

Private Sub btnAdicionar_Click()
On Error GoTo tratar_erro

ProcAdicionar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_carregarinfest_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cliente_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_pedido_todos_Click()
On Error GoTo tratar_erro

GridItens.rows = 1
If Chk_pedido_todos.Value = 1 Then
    With Txt_pedido_cliente
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
Else
    With Txt_pedido_cliente
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_prazo_todos_Click()
On Error GoTo tratar_erro

GridItens.rows = 1
If Chk_prazo_todos.Value = 1 Then
    With Txt_dias
        .Locked = False
        .TabStop = True
        If Vendas_Proposta = True Then .SetFocus
    End With
    With Cmb_prazo
        .Enabled = True
        If Vendas_Proposta = False Then .SetFocus
    End With
Else
    With Txt_dias
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    With Cmb_prazo
        .Value = Date
        .Enabled = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_prazo_Change()
On Error GoTo tratar_erro

GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1
With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Vendas = 'True'", True
        ElseIf cmbfiltrarpor = "Cliente" Then
                Set TBClientes = CreateObject("adodb.recordset")
                TBClientes.Open "Select IDCliente, NomeRazao from Clientes where NomeRazao <> 'Null' order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    Do While TBClientes.EOF = False
                        .AddItem TBClientes!NomeRazao
                        .ItemData(.NewIndex) = TBClientes!IDCliente
                        TBClientes.MoveNext
                    Loop
                    If Vendas_Programacao = True Then
                        cmbfamilia.Text = frmVendas_programacao.txtCliente
                    Else
                        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
                            If .txtIDcliente <> "" And .txtIDcliente <> "0" Then cmbfamilia.Text = .txtCliente
                        End With
                    End If
                End If
                TBClientes.Close
            Else
                Set TBFamilia = CreateObject("adodb.recordset")
                TBFamilia.Open "Select * from Clientes_grupos where Texto <> 'Null' order by Texto", Conexao, adOpenKeyset, adLockOptimistic
                If TBFamilia.EOF = False Then
                    Do While TBFamilia.EOF = False
                        .AddItem TBFamilia!Texto
                        .ItemData(.NewIndex) = TBFamilia!ID
                        TBFamilia.MoveNext
                    Loop
                End If
                TBFamilia.Close
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Chk_pedido_todos.Value = 1 And Txt_pedido_cliente = "" Then
    NomeCampo = "o pedido do cliente"
    ProcVerificaAcao
    Txt_pedido_cliente.SetFocus
    Exit Sub
End If
If Vendas_Proposta = True And Chk_prazo_todos.Value = 1 And Txt_dias = "" Then
    NomeCampo = "o prazo em dias"
    ProcVerificaAcao
    Txt_dias.SetFocus
    Exit Sub
End If

If Proposta_Servicos = True Or PI_Servicos = True Then TipoProduto = "P.Tipo = 'S'" Else TipoProduto = "P.tipo = 'P'"
CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Descricaotecnica, P.Unidade, P.Unidade_com, P.classe, P.Producao, P.Compras, P.Vendas"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Projproduto_clientes PC ON PC.Codproduto = P.Codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
If Cmb_ordenar = "Código interno" Then Ordenar = "P.desenho" Else Ordenar = "P.Descricao"

TextoFiltroCliente = ""
If Chk_cliente.Value = 1 Then
    If Vendas_Programacao = True Then
        TextoFiltroCliente = "and PC.Idcliente = " & frmVendas_programacao.txtID_cli
    Else
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            TextoFiltroCliente = "and PC.Idcliente = " & .txtIDcliente
        End With
    End If
End If
TextoFiltroPadrao = "P.Vendas = 'True' and " & TipoProduto & " and P.bloqueado = 'False' and P.DtValidacao IS NOT NULL " & TextoFiltroCliente & " group by " & CamposFiltro & " order by " & Ordenar

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Cliente" Then
        StrSqlLocProdVendas = INNERJOINTEXTO & " where PC.IDCliente = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "Grupo do cliente" Then
            StrSqlLocProdVendas = INNERJOINTEXTO & " where IA.IDGrupo = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Família" Then
                StrSqlLocProdVendas = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                    Select Case cmbfiltrarpor
                        Case "Comprimento": TextoFiltro = "P.Comprimento"
                        Case "Largura": TextoFiltro = "P.Largura"
                        Case "Espessura": TextoFiltro = "P.Espessura"
                    End Select
                    valor = txtTexto
                    NovoValor = Replace(valor, ",", ".")
                    StrSqlLocProdVendas = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
                Else
                    Select Case cmbfiltrarpor
                        Case "Código interno": TextoFiltro = "P.desenho"
                        Case "Código de referência": TextoFiltro = "IA.N_referencia"
                        Case "Número do desenho": TextoFiltro = "IA.desenho"
                        Case "Descrição": TextoFiltro = "P.descricao"
                        Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                        Case "Dureza": TextoFiltro = "P.Dureza"
                        Case "Part number": TextoFiltro = "PFAB.Part_number"
                    End Select
                    StrSqlLocProdVendas = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLocProdVendas = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_pesquisar_estrutura_Click()
On Error GoTo tratar_erro

ProcCarregaEstrutura

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
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
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
On Error GoTo tratar_erro

With FlexGrid
    If Col = 12 Then
        If Vendas_Proposta = False Then
            vNewValue = Format(vNewValue, "DD/MM/YY")
        Else
            vNewValue = Format(vNewValue, "###,##0")
        End If
    End If
    If Col = 13 Then vNewValue = Format(vNewValue, "###,##0.0000")
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridItens()
On Error GoTo tratar_erro

With GridItens

    .AllowUserPaste = cellTextOnly
    .AllowUserResizing = True
    .ExtendLastCol = True
    .BoldFixedCell = False
    .DisplayDateTimeMask = True
    .DisplayFocusRect = False
    .SelectionMode = cellSelectionFree
    .Cols = 13
    .DrawMode = cellOwnerDraw
    .Column(0).Width = 20
    .Column(0).Alignment = cellCenterCenter
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    
    .Cell(0, 1).ForeColor = vbRed
    .Cell(0, 1).Text = ""
    .Column(1).CellType = cellCheckBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Width = 20
    
    .Cell(0, 2).ForeColor = vbRed
    .Cell(0, 2).Text = "Cód. interno"
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellLeftCenter
    .Column(2).Locked = True
    .Column(2).Width = 80

    .Cell(0, 3).ForeColor = vbRed
    .Cell(0, 3).Text = "Referência"
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellLeftCenter
    .Column(3).Locked = True
    .Column(3).Width = 150
    
    .Cell(0, 4).ForeColor = vbRed
    .Cell(0, 4).Text = "Descrição comercial"
    .Column(4).CellType = cellTextBox
    .Column(4).Alignment = cellLeftCenter
    .Column(4).Locked = True
    .Column(4).Width = 400
    
    .Cell(0, 5).ForeColor = vbRed
    .Cell(0, 5).Text = "Un com"
    .Column(5).CellType = cellTextBox
    .Column(5).Alignment = cellCenterCenter
    .Column(5).Locked = True
    .Column(5).Width = 50

    .Cell(0, 6).ForeColor = vbRed
    .Cell(0, 6).Text = "Saldo est."
    .Column(6).CellType = cellTextBox
    .Column(6).Alignment = cellCenterCenter
    .Column(6).Locked = True
    If Prod = True Then
    .Column(6).Width = 65
    .Column(7).Width = 68
    .Column(8).Width = 68
    Chk_carregarinfest.Visible = True
    Else
    .Column(6).Width = 0
    .Column(7).Width = 0
    .Column(8).Width = 0
    Chk_carregarinfest.Visible = False
    End If
    
    .Cell(0, 7).ForeColor = vbRed
    .Cell(0, 7).Text = "Empenhado"
    .Column(7).CellType = cellTextBox
    .Column(7).Alignment = cellCenterCenter
    .Column(7).Locked = True

    
    .Cell(0, 8).ForeColor = vbRed
    .Cell(0, 8).Text = "Diponível"
    .Column(8).CellType = cellTextBox
    .Column(8).Alignment = cellCenterCenter
    .Column(8).Locked = True


    .Cell(0, 9).ForeColor = vbRed
    .Cell(0, 9).Text = "Quantidade"
    .Column(9).CellType = cellTextBox
    .Column(9).Alignment = cellCenterCenter
    .Column(9).Width = 68

    .Cell(0, 10).ForeColor = vbRed
    .Cell(0, 10).Text = "Prazo"
    
    If Vendas_PI = True Then
    .Column(10).CellType = cellDate
    Else
    .Column(10).CellType = cellTextBox
    End If
    
    .Column(10).Alignment = cellCenterCenter
    .Column(10).Width = 85
    
    .Cell(0, 11).ForeColor = vbRed
    .Cell(0, 11).Text = "Codproduto"
    .Column(11).CellType = cellTextBox
    .Column(11).Alignment = cellCenterCenter
    .Column(11).Width = 0
    
    .Cell(0, 12).ForeColor = vbRed
    .Cell(0, 12).Text = "Pedido cliente"
    .Column(12).CellType = cellTextBox
    .Column(12).Alignment = cellCenterCenter
    
    If Vendas_PI = True Then
    .Column(12).Width = 55
    Else
    .Column(12).Width = 0
    End If
    

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGrid_CellClick(ByVal Row As Long, ByVal Col As Long, Shift As Integer)
On Error GoTo tratar_erro

FlexGrid_Desenho = FlexGrid.CellText(Row, 2)

Permitido = True
If Col = 1 And FlexGrid.CellChecked(Row, 1) = True Then
'    If FlexGrid.CellText(Row, 12) = "" Then
'        usMsgbox ("Informe o " & IIf(Vendas_Proposta = True, "prazo (dias)", "prazo") & " antes de adicionar."), vbExclamation, "CAPRIND v5.0"
'        FlexGrid.CellChecked(Row, 1) = False
'        Exit Sub
'    End If
'    If FlexGrid.CellText(Row, 13) = "" Then
'        usMsgbox ("Informe a quantidade antes de adicionar."), vbExclamation, "CAPRIND v5.0"
'        FlexGrid.CellChecked(Row, 1) = False
'        Exit Sub
'    End If
'    If FlexGrid.CellText(Row, 14) = "" And Vendas_PI = True Then
'        usMsgbox ("Informe o pedido do cliente antes de adicionar."), vbExclamation, "CAPRIND v5.0"
'        FlexGrid.CellChecked(Row, 1) = False
'        Exit Sub
'    End If

    'Verifica se o produto/serviço pertence ao cliente
    If Vendas_Programacao = True Then
        With frmVendas_programacao
            IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            IDCliente = .txtID_cli
            Cliente = .txtCliente
        End With
    Else
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            IDCliente = .txtIDcliente
            Cliente = .txtCliente
        End With
    End If
    If IDCliente <> 0 Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & FlexGrid.CellText(Row, 0) & " and IDCliente <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from Projproduto_clientes where Codproduto = " & FlexGrid.CellText(Row, 0) & " and IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = True Then
                If PI_Servicos = True Or Proposta_Servicos = True Then NomeCampo = "serviço" Else NomeCampo = "Produto"
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from empresa where codigo = " & IDempresa & " and Bloquear_produtos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    If USMsgBox("Este " & NomeCampo & " não pertence ao cliente " & Cliente & ", deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                        FlexGrid.CellChecked(Row, 1) = False
                    End If
                Else
                    USMsgBox ("Este " & NomeCampo & " não pertence ao cliente " & Cliente & "."), vbExclamation, "CAPRIND v5.0"
                    FlexGrid.CellChecked(Row, 1) = False
                End If
                TBCiclo.Close
            End If
            TBProduto.Close
        End If
        TBAbrir.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGrid_Click()

'usMsgbox "teste"

End Sub

Private Sub FlexGrid_ColumnClick(ByVal Col As Long)
On Error GoTo tratar_erro

If Col = 1 And Permitido = False Then
    With FlexGrid
        For InitFor = 0 To (.rows + 1)
            If .CellChecked(InitFor, 1) = True Then
                .CellChecked(InitFor, 1) = False
            Else
'                If FlexGrid.CellText(InitFor, 12) = "" Then GoTo Proximo
'                If FlexGrid.CellText(InitFor, 13) = "" Then GoTo Proximo
'                If FlexGrid.CellText(InitFor, 14) = "" And Vendas_PI = True Then GoTo Proximo
            
                'Verifica se o produto/serviço pertence ao cliente
                If Vendas_Programacao = True Then
                    With frmVendas_programacao
                        IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
                        IDCliente = .txtID_cli
                        Cliente = .txtCliente
                    End With
                Else
                    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
                        IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
                        IDCliente = .txtIDcliente
                        Cliente = .txtCliente
                    End With
                End If
                If IDCliente <> 0 Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & FlexGrid.CellText(InitFor, 0) & " and IDCliente <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select * from Projproduto_clientes where Codproduto = " & FlexGrid.CellText(InitFor, 0) & " and IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = True Then
                            If PI_Servicos = True Or Proposta_Servicos = True Then NomeCampo = "serviço" Else NomeCampo = "Produto"
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from empresa where codigo = " & IDempresa & " and Bloquear_produtos = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                TBCiclo.Close
                                TBProduto.Close
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                            TBCiclo.Close
                        End If
                        TBProduto.Close
                    End If
                    TBAbrir.Close
                End If
            
            .CellChecked(InitFor, 1) = True
Proximo:
            End If
        Next InitFor
    End With
End If
Permitido = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

'If 'USToolBar1.ButtonState(1) = 5 Then Exit Sub
If SSTab1.Tab = 0 Then
    Select Case KeyCode
        Case vbKeyReturn: If GridItens.Visible = False Then ListView1_DblClick
        Case vbKeyF2: ProcFiltrar
        Case vbKeyF3: If GridItens.Visible = True Then ProcAdicionar
        'Case vbKeyF1: ProcAjuda
        Case vbKeyEscape: ProcSair
    End Select
Else
    Select Case KeyCode
        'Case vbKeyF1: ProcAjuda
        Case vbKeyEscape: ProcSair
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

FlexGrid_Desenho = ""
'ProcCarregaToolBar1 Me, 17445, 6, True
ProcCarregaToolBar2 Me, 17445, 3, True
ProcAjustaGridItens

Chk_cliente.Caption = "Filtrar produtos do cliente"
If Vendas_Programacao = True Then
    Caption = "Administrativo - Vendas - Programação - Localizar produtos"
    IDConta = frmVendas_programacao.Cmb_empresa.ItemData(frmVendas_programacao.Cmb_empresa.ListIndex)
    'USToolBar1.ButtonState(2) = 5
ElseIf PI_Produtos = True Then
        If Vendas_PI = True Then
            Caption = "Administrativo - Vendas - Pedido interno - Localizar produtos"
            IDConta = frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex)

            'Adicionar coluna ao flexGrid
            If frmVendas_PI.txtid_produto = 0 Then
                GridItens.Visible = True
                ListView1.Visible = False
                'FlexGrid.ColWidth(4) = 3035
                'FlexGrid.AddColumn "Prazo", 899, AlignToCenterCenter, TypeDate
            Else
                'USToolBar1.ButtonState(2) = 5
            End If
        Else
            IDConta = frmVendas_proposta.Cmb_empresa.ItemData(frmVendas_proposta.Cmb_empresa.ListIndex)
            'Adicionar coluna ao flexGrid
            If frmVendas_proposta.txtid_produto = 0 Then
                GridItens.Visible = True
                ListView1.Visible = False
            Else
                'USToolBar1.ButtonState(2) = 5
            End If
        End If
    ElseIf PI_Servicos = True Then
        With SSTab1
            If Vendas_PI = True Then
                Caption = "Administrativo - Vendas - Pedido interno - Localizar serviços"
                IDConta = frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex)
                If frmVendas_PI.txtid_servico = 0 Then
                    'Adicionar coluna ao flexGrid
                    GridItens.Visible = True
                    ListView1.Visible = False
                Else
                    'USToolBar1.ButtonState(2) = 5
                End If
            Else
                IDConta = frmVendas_proposta.Cmb_empresa.ItemData(frmVendas_proposta.Cmb_empresa.ListIndex)
                'Adicionar coluna ao flexGrid
                If frmVendas_proposta.txtid_servico = 0 Then
                    GridItens.Visible = True
                    ListView1.Visible = False
                Else
                    'USToolBar1.ButtonState(2) = 5
                End If
            End If
            .TabVisible(1) = False
            .TabsPerRow = 1
            .TabCaption(0) = "Lista de serviços"
        End With
        Chk_cliente.Caption = "Filtrar serviços do cliente"
End If
If Vendas_Proposta = True Then
    Chk_pedido_todos.Visible = False
    Txt_pedido_cliente.Visible = False
    With Txt_dias
        .Visible = True
    End With
    Cmb_prazo.Visible = False
End If
'USToolBar1.Refresh

SSTab1.Tab = 0

Cmb_prazo.Value = Date
ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, IDempresa, "Produtos/Serviços", "V", True
If Permitido = False Then cmbfiltrarpor = "Código interno"

Cmb_ordenar = "Código interno"
cmbVersao_pesquisar_estrutura = "A"

ProcRemoveObjetosResize Me
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
GridItens.rows = 1
If StrSqlLocProdVendas = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdVendas, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro
Dim L As Long

GridItens.rows = 1
ListView1.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)

TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

Contador = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador2 = 0

Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    If GridItens.Visible = True Then
        Frame1.Enabled = False
        Frame9.Enabled = False
        With USToolBar2
            .ButtonState(1) = 5
            .ButtonState(2) = 5
            .Refresh
        End With
        
        With GridItens
        L = Contador
            .AddItem TBLocalizar_produto_padrao!Codproduto
            ' contador
            .Cell(L, 1).Text = ""
            
            .Cell(L, 2).Text = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
                                    
            If Vendas_Programacao = True Then
                IDCliente = frmVendas_programacao.txtID_cli
            Else
               IDCliente = IIf(Vendas_PI = True, frmVendas_PI.txtIDcliente, frmVendas_proposta.txtIDcliente)
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & TBLocalizar_produto_padrao!Codproduto & " and IA.ID_cliente_forn = " & IDCliente & " and IA.Tipo = 'C' and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & TBLocalizar_produto_padrao!Codproduto & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            End If
            If TBAbrir.EOF = False Then
                .Cell(L, 3).Text = TBAbrir!N_referencia
            End If
            TBAbrir.Close
            
           ' .Cell(L, 4).Text = IIf(IsNull(TBLocalizar_produto_padrao!descricaotecnica), "", TBLocalizar_produto_padrao!descricaotecnica)
            .Cell(L, 4).Text = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
           ' .Cell(L, 6).Text = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
            .Cell(L, 5).Text = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
           ' .Cell(L, 8).Text = IIf(IsNull(TBLocalizar_produto_padrao!Classe), "", TBLocalizar_produto_padrao!Classe)
            
 '==========================================================
 ' Se for pra carregar valores de estoque
 '==========================================================
            If Chk_carregarinfest.Value = 1 Then
                Valor1 = FunVerificaQtdeEstoque(TBLocalizar_produto_padrao!Desenho, IDempresa, "")
                .Cell(L, 6).Text = Valor1
                
                Valor2 = FunVerificaQtdeEmpenhoEstVenda(TBLocalizar_produto_padrao!Desenho, IDempresa)
                TTE = FunVerificaQtdeEmpenhoEst(TBLocalizar_produto_padrao!Desenho, IDempresa)
                .Cell(L, 7).Text = Valor2 + TTE
                
                Valor_Cofins_Prod = Valor1 - (Valor2 + TTE)
                .Cell(L, 8).Text = Valor_Cofins_Prod
            End If
            
            If Chk_pedido_todos.Value = 1 Then .Cell(L, 9).Text = Txt_pedido_cliente
            If Chk_prazo_todos.Value = 1 Then .Cell(L, 10).Text = IIf(Vendas_Proposta = True, Txt_dias, Format(Cmb_prazo, "dd/mm/yy"))
            
            .Cell(L, 11).Text = TBLocalizar_produto_padrao!Codproduto
            
        End With
        Contador = Contador + 1
    Else
        With ListView1.ListItems
            .Add , , TBLocalizar_produto_padrao!Codproduto
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!descricaotecnica), "", TBLocalizar_produto_padrao!descricaotecnica)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!Classe), "", TBLocalizar_produto_padrao!Classe)
            
            If Chk_carregarinfest.Value = 1 Then
                Valor1 = FunVerificaQtdeEstoque(TBLocalizar_produto_padrao!Desenho, IDempresa, "")
                .Item(.Count).SubItems(7) = Valor1
                        
                Valor2 = FunVerificaQtdeEmpenhoEstVenda(TBLocalizar_produto_padrao!Desenho, IDempresa)
                TTE = FunVerificaQtdeEmpenhoEst(TBLocalizar_produto_padrao!Desenho, IDempresa)
                .Item(.Count).SubItems(8) = Valor2 + TTE
                Valor_Cofins_Prod = Valor1 - (Valor2 + TTE)
                .Item(.Count).SubItems(9) = Valor_Cofins_Prod
            End If
        End With
    End If
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador2 = Contador2 + 1
    PBLista.Value = Contador2
Loop
lblRegistros.Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If
If GridItens.Visible = True Then
    'FlexGrid.Redraw = True
'    With USToolBar1
'        .ButtonState(1) = 0
'        .ButtonState(2) = 0
'        .ButtonState(4) = 0
'        .ButtonState(5) = 0
'        .Refresh
'    End With
    'Frame3.Enabled = True
 '   Frame5.Enabled = True
    Frame1.Enabled = True
    Frame9.Enabled = True
    With USToolBar2
        .ButtonState(1) = 0
        .ButtonState(2) = 0
        .Refresh
    End With
Else
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExibePaginaOLD(Pagina)
On Error GoTo tratar_erro
Dim L As Long

GridItens.rows = 1
ListView1.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)

TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

Contador = 1
PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
Contador2 = 0

Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    If GridItens.Visible = True Then
'        With USToolBar1
'            .ButtonState(1) = 5
'            .ButtonState(2) = 5
'            .ButtonState(4) = 5
'            .ButtonState(5) = 5
'            .Refresh
'        End With
        'Frame3.Enabled = False
        'Frame5.Enabled = False
        Frame1.Enabled = False
        Frame9.Enabled = False
        With USToolBar2
            .ButtonState(1) = 5
            .ButtonState(2) = 5
            .Refresh
        End With
        
        With FlexGrid
            L = .AddItem(TBLocalizar_produto_padrao!Codproduto)
            .RowData(L) = Contador
            .CellText(L, 2) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
                                    
            If Vendas_Programacao = True Then
                IDCliente = frmVendas_programacao.txtID_cli
            Else
               IDCliente = IIf(Vendas_PI = True, frmVendas_PI.txtIDcliente, frmVendas_proposta.txtIDcliente)
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & TBLocalizar_produto_padrao!Codproduto & " and IA.ID_cliente_forn = " & IDCliente & " and IA.Tipo = 'C' and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & TBLocalizar_produto_padrao!Codproduto & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            End If
            If TBAbrir.EOF = False Then
                .CellText(L, 3) = TBAbrir!N_referencia
            End If
            TBAbrir.Close
            
            .CellText(L, 4) = IIf(IsNull(TBLocalizar_produto_padrao!descricaotecnica), "", TBLocalizar_produto_padrao!descricaotecnica)
            .CellText(L, 5) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
            .CellText(L, 6) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
            .CellText(L, 7) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .CellText(L, 8) = IIf(IsNull(TBLocalizar_produto_padrao!Classe), "", TBLocalizar_produto_padrao!Classe)
            
            If Chk_carregarinfest.Value = 1 Then
                Valor1 = FunVerificaQtdeEstoque(TBLocalizar_produto_padrao!Desenho, IDempresa, "")
                .CellText(L, 9) = Valor1
                
                Valor2 = FunVerificaQtdeEmpenhoEstVenda(TBLocalizar_produto_padrao!Desenho, IDempresa)
                TTE = FunVerificaQtdeEmpenhoEst(TBLocalizar_produto_padrao!Desenho, IDempresa)
                .CellText(L, 10) = Valor2 + TTE
                
                Valor_Cofins_Prod = Valor1 - (Valor2 + TTE)
                .CellText(L, 11) = Valor_Cofins_Prod
            End If
            
            If Chk_pedido_todos.Value = 1 Then .CellText(L, 14) = Txt_pedido_cliente
            If Chk_prazo_todos.Value = 1 Then .CellText(L, 12) = IIf(Vendas_Proposta = True, Txt_dias, Format(Cmb_prazo, "dd/mm/yy"))
            
        End With
        Contador = Contador + 1
    Else
        With ListView1.ListItems
            .Add , , TBLocalizar_produto_padrao!Codproduto
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Desenho), "", TBLocalizar_produto_padrao!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!descricaotecnica), "", TBLocalizar_produto_padrao!descricaotecnica)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_produto_padrao!Classe), "", TBLocalizar_produto_padrao!Classe)
            
            If Chk_carregarinfest.Value = 1 Then
                Valor1 = FunVerificaQtdeEstoque(TBLocalizar_produto_padrao!Desenho, IDempresa, "")
                .Item(.Count).SubItems(7) = Valor1
                        
                Valor2 = FunVerificaQtdeEmpenhoEstVenda(TBLocalizar_produto_padrao!Desenho, IDempresa)
                TTE = FunVerificaQtdeEmpenhoEst(TBLocalizar_produto_padrao!Desenho, IDempresa)
                .Item(.Count).SubItems(8) = Valor2 + TTE
                Valor_Cofins_Prod = Valor1 - (Valor2 + TTE)
                .Item(.Count).SubItems(9) = Valor_Cofins_Prod
            End If
        End With
    End If
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador2 = Contador2 + 1
    PBLista.Value = Contador2
Loop
lblRegistros.Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If
If FlexGrid.Visible = True Then
    FlexGrid.Redraw = True
    With USToolBar1
        .ButtonState(1) = 0
        .ButtonState(2) = 0
        .ButtonState(4) = 0
        .ButtonState(5) = 0
        .Refresh
    End With
    Frame3.Enabled = True
    Frame5.Enabled = True
    Frame1.Enabled = True
    Frame9.Enabled = True
    With USToolBar2
        .ButtonState(1) = 0
        .ButtonState(2) = 0
        .Refresh
    End With
Else
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridItens_CellChange(ByVal Row As Long, ByVal Col As Long)

With GridItens
vRow = .ActiveCell.Row

If .Cell(vRow, 1).Text = "1" Then
.Cell(vRow, 9).Locked = False
.Cell(vRow, 10).Locked = False
.Cell(vRow, 9).SetFocus
Else
If vRow > 0 Then
.Cell(vRow, 9).Text = ""
.Cell(vRow, 10).Text = ""
End If
.Cell(vRow, 9).Locked = True
.Cell(vRow, 10).Locked = True
End If

End With

End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If (ListView1.Visible = True And ListView1.ListItems.Count = 0) Or (GridItens.Visible = True And FlexGrid_Desenho = "") Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        If ListView1.Visible = True Then ListView1.SetFocus
        PBLista.Visible = True
    Case 1:
        PBLista.Visible = False
        cmbVersao_pesquisar_estrutura.SetFocus
        ProcCarregaEstrutura
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEstrutura()
On Error GoTo tratar_erro
''ReDim arrNodes(2000)

If ListView1.ListItems.Count = 0 And FlexGrid_Desenho = "" Then Exit Sub
Call m_Tree.Nodes.Clear
Grid1.rows = 1

m_Row = 1
m_Col = 1

Contador1 = -1
Set TBLISTA = CreateObject("adodb.recordset")
If GridItens.Visible = True Then TextoFiltro = FlexGrid_Desenho Else TextoFiltro = ListView1.SelectedItem.ListSubItems(1)
TBLISTA.Open "Select * from Projproduto where Desenho = '" & TextoFiltro & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    CodRef = ""
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        CodRef = TBFI!N_referencia
    End If
    TBFI.Close
    
    DataValidacao = ""
    RespValidacao = ""
    If TBLISTA!SubTipoItem <> 0 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Projconjunto_desc_versao where codproduto = " & TBLISTA!Codproduto & " and Versao = '" & cmbVersao_pesquisar_estrutura & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            DataValidacao = IIf(IsNull(TBFI!DtValidacao), "", TBFI!DtValidacao)
            RespValidacao = IIf(IsNull(TBFI!RespValidacao), "", TBFI!RespValidacao)
        End If
    End If
    Contador1 = Contador1 + 1
    arrNodes(Contador1).Level = 0
    arrNodes(Contador1).Text = TBLISTA!Desenho & vbTab & "" & vbTab & TBLISTA!Codproduto & vbTab & CodRef & vbTab & TBLISTA!Descricao & vbTab & TBLISTA!Unidade & vbTab & cmbVersao_pesquisar_estrutura & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao
    
    Codproduto = TBLISTA!Codproduto
    
    ProcNivel2Estrutura frmVendas_ListaProduto, cmbVersao_pesquisar_estrutura, False, False, True, False

    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 20
        .rows = m_Row
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Pos."
        .Cell(0, 3).Text = "ID"
        .Cell(0, 4).Text = "Cód. de ref."
        .Cell(0, 5).Text = "Descrição"
        .Cell(0, 6).Text = "Un."
        .Cell(0, 7).Text = "Ver."
        .Cell(0, 8).Text = "Kg/un"
        .Cell(0, 9).Text = "Un/kg"
        .Cell(0, 10).Text = "Dim/mm"
        .Cell(0, 11).Text = "Kg/pç"
        .Cell(0, 12).Text = "Qtde."
        .Cell(0, 13).Text = "Peso total"
        .Cell(0, 14).Text = "Vlr. custo"
        .Cell(0, 15).Text = "Dt. validação"
        .Cell(0, 16).Text = "Resp. validação"
        .Cell(0, 17).Text = "ID estr."
        .Cell(0, 18).Text = "Part number"
        .Cell(0, 19).Text = "Observações"
        .Range(0, 1, 0, 19).Alignment = cellCenterCenter
        .Column(1).Width = 200
        .Column(2).Width = 30
        .Column(3).Width = 0
        .Column(4).Width = 80
        .Column(5).Width = 300
        .Column(6).Width = 40
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Width = 40
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Width = 100
        .Column(8).Alignment = cellRightCenter
        .Column(9).Width = 40
        .Column(10).Width = 100
        .Column(10).Alignment = cellRightCenter
        .Column(11).Width = 100
        .Column(11).Alignment = cellRightCenter
        .Column(12).Width = 100
        .Column(12).Alignment = cellRightCenter
        .Column(13).Width = 100
        .Column(13).Alignment = cellRightCenter
        .Column(14).Width = 100
        .Column(14).Alignment = cellRightCenter
        .Column(15).Width = 120
        .Column(16).Width = 100
        .Column(17).Width = 0
        .Column(18).Width = 150
        .Column(19).Width = 400
        
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text
        
        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            .AddItem arrNodes(intIndex).Text
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_dias_Change()
On Error GoTo tratar_erro

GridItens.rows = 1
If Txt_dias <> "" Then
    VerifNumero = Txt_dias
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_dias = ""
        Txt_dias.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_dias_LostFocus()
On Error GoTo tratar_erro

Txt_dias = Format(Txt_dias, "###,##0")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pedido_cliente_Change()
On Error GoTo tratar_erro

GridItens.rows = 1

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

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
GridItens.rows = 1
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAdicionar()
On Error GoTo tratar_erro

'FlexGrid_Click
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido2 = False
If PI_Produtos = True Then
    MsgTexto = "produto"
    MsgTexto1 = "Produto"
Else
    MsgTexto = "serviço"
    MsgTexto1 = "Serviço"
End If

Contador = 0
Permitido = False
With GridItens
    For InitFor = 1 To (.rows)
        If .Cell(Contador, 1).Text = "1" Then
        
            If .Cell(Contador, 9).Text = "" Then
               USMsgBox "Informe a quantidade por favor", vbCritical, "CAPRIND v5.0"
             .Cell(Contador, 9).SetFocus
             Exit Sub
            End If
        
            If .Cell(Contador, 10).Text = "" Then
               USMsgBox "Informe o prazo de entrega por favor", vbCritical, "CAPRIND v5.0"
             .Cell(Contador, 10).SetFocus
             Exit Sub
            End If
        End If
        Contador = Contador + 1
    Next InitFor
End With

Contador = 0
With GridItens
    For InitFor = 1 To (.rows)
        If .Cell(Contador, 1).Text = "1" Then
            If Permitido2 = False Then
                If USMsgBox("Deseja realmente adicionar este(s) " & MsgTexto & "(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If

            Permitido2 = True
            procCriarItem .Cell(Contador, 12).Text, .Cell(Contador, 11).Text, .Cell(Contador, 9).Text, .Cell(Contador, 10).Text
        End If
        Contador = Contador + 1
    Next InitFor
End With

If Permitido2 = False Then
    USMsgBox ("Informe o(s) " & MsgTexto & "(s) antes de adicionar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    USMsgBox (MsgTexto1 & " adicionado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        If PI_Servicos = True Then
            .ProcAtualizalistaServicos (IIf(ReturnNumbersOnly(Left(.lblPaginas2.Caption, Len(.lblPaginas2.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas2.Caption, Len(.lblPaginas2.Caption) - 5))))
            If Vendas_Proposta = True Then .Novo_Proposta2 = False Else .Novo_PI2 = False
        Else
            .ProcAtualizalistaProdutos (IIf(ReturnNumbersOnly(Left(.lblPaginas1.Caption, Len(.lblPaginas1.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas1.Caption, Len(.lblPaginas1.Caption) - 5))))
            If Vendas_Proposta = True Then .Novo_Proposta1 = False Else .Novo_PI1 = False
        End If
        .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCriarItem(PedidoCli_adicionar As String, Codprod_adicionar As Long, Quantidade_adicionar As Double, Prazo_adicionar As String)
On Error GoTo tratar_erro

If Vendas_Programacao = False Then
    Set TBCotacao = CreateObject("adodb.recordset")
    TBCotacao.Open "Select * from vendas_carteira", Conexao, adOpenKeyset, adLockOptimistic
    TBCotacao.AddNew
    If Vendas_PI = True Then
        TBCotacao!Liberacao = "VENDIDA"
        TBCotacao!Datavendas = frmVendas_PI.txtDatavendas_PI
    Else
        TBCotacao!Liberacao = "ABERTA EM ANALISE"
    End If
    TBCotacao!Tem_ordem = False
    TBCotacao!IDAnalise = 0
    TBCotacao!quantidade = Quantidade_adicionar
    TBCotacao!PCCliente = IIf(PedidoCli_adicionar = "", Null, PedidoCli_adicionar)
    If Vendas_PI = True Then
        TBCotacao!PrazoFinal = Prazo_adicionar
        TBCotacao!Prazo_original = Prazo_adicionar
    Else
        TBCotacao!prazofinaldias = IIf(IsNumeric(Prazo_adicionar), Prazo_adicionar, 0)
    End If
    
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where codproduto = " & Codprod_adicionar, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        TBCotacao!Desenho = TBItem!Desenho
        TBCotacao!Rev_codinterno = TBItem!RevDesenho
        TBCotacao!Descricao = TBItem!descricaotecnica
        TBCotacao!Qtde_produzir = TBCotacao!quantidade / FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)
        TBCotacao!Antecipacao_fat = False
        TBCotacao!Faturamento_parcial = False
        TBCotacao!Comprimento = TBItem!Comprimento
        TBCotacao!Largura = TBItem!Largura
        TBCotacao!Espessura = TBItem!Espessura
        TBCotacao!descricao_tecnica = TBItem!Descricao
        TBCotacao!Unidade = TBItem!Unidade
        TBCotacao!Unidade_com = TBItem!Unidade_com
        TBCotacao!Familia = TBItem!Classe
        If PI_Produtos = True Then TBCotacao!Tipo = "P" Else TBCotacao!Tipo = "S"
        TBCotacao!retorno = False
        
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & Codprod_adicionar & " and IA.ID_cliente_forn = " & .txtIDcliente & " and IA.Tipo = 'C' and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.codproduto = " & Codprod_adicionar & " and (IA.ID_cliente_forn = 0 or IA.ID_cliente_forn IS NULL) and IA.N_Referencia is not null group by IA.n_referencia", Conexao, adOpenKeyset, adLockOptimistic
            End If
            If TBAbrir.EOF = False Then
                TBCotacao!N_referencia = TBAbrir!N_referencia
            End If
            TBAbrir.Close
                
            TBCotacao!Cotacao = .txtId
            valor = IIf(.Txt_valor_moeda = "", 1, .Txt_valor_moeda)
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & TBItem!Codproduto & " and idcliente = " & .txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If .txttipocliente <> "JR" And .txttipocliente <> "FR" Then
                    TBCotacao!preco_unitario = IIf(IsNull(TBAbrir!PConsumo), 0, Format((TBAbrir!PConsumo / FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)) / valor, "###,##0.0000000000"))
                Else
                    TBCotacao!preco_unitario = IIf(IsNull(TBAbrir!PRevenda), 0, Format((TBAbrir!PRevenda / FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)) / valor, "###,##0.0000000000"))
                End If
                TBCotacao!ID_CF = IIf(TBAbrir!ID_CF = "", Null, TBAbrir!ID_CF)
            Else
                If .txttipocliente <> "JR" And .txttipocliente <> "FR" Then
                    TBCotacao!preco_unitario = IIf(IsNull(TBItem!PConsumo), "", Format((TBItem!PConsumo / FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)) / valor, "###,##0.0000000000"))
                Else
                    TBCotacao!preco_unitario = IIf(IsNull(TBItem!PRevenda), "", Format((TBItem!PRevenda / FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)) / valor, "###,##0.0000000000"))
                End If
            End If
            TBAbrir.Close
            
            If IsNull(TBCotacao!ID_CF) = True Or TBCotacao!ID_CF = "" Then TBCotacao!ID_CF = TBItem!ID_CF
            TBCotacao!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP1), 0, TBItem!ID_CFOP1)
            
            If TBCotacao!ID_CFOP = 0 Then
                'Verifica CFOP vinculada ao cliente
                Set TBOSC = CreateObject("adodb.recordset")
                TBOSC.Open "Select IDCFOP FROM Clientes_DadosComerciais where IDCliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente) & " and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and IDCFOP IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBOSC.EOF = False Then
                    TBCotacao!ID_CFOP = TBOSC!IDCFOP
                End If
                TBOSC.Close
            Else
            '========================
            ' Busca CST do ICMS
            '========================
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select * FROM tbl_NaturezaOperacao_CST where ID_CFOP = " & TBCotacao!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.EOF = False Then
                
                If TBAliquota.RecordCount > 1 Then
                frmVendas_PI_CST.Show 1
                Else
                TBCotacao!txt_CST = TBAliquota!CST_ICMS
                End If
                End If
                'TBAliquota.Close
            End If
            
            If IsNull(TBCotacao!ID_CF) = False And TBCotacao!ID_CF <> "" And PI_Produtos = True Then
                ProcValorImposto .txtCotacao, TBCotacao!ID_CF, IIf(.txtIDcliente = "", 0, .txtIDcliente), .txtCliente, .txtuf, .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), False, IIf(IsNull(TBItem!ID_CFOP1), 0, TBItem!ID_CFOP1), IIf(Vendas_PI = True, frmVendas_PI.RegimeEmpresa_PI, frmVendas_proposta.RegimeEmpresa_Proposta)
                ProcControleImposto TBCotacao!ID_CFOP, IIf(.txtIDcliente = "", 0, .txtIDcliente)
                If TemIPI = "SIM" Then
                    TBCotacao!int_IPI = IntIPI
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select * from Clientes_Impostos where IDCliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente) & " and ID_CF = " & TBCotacao!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        VlrIPI = TBCotacao!preco_unitario
                        If TBFIltro!PorcentagemIPI <> 0 Then VlrIPI = VlrIPI / TBFIltro!PorcentagemIPI
                        VlrIPI = (VlrIPI - TBCotacao!preco_unitario) * TBCotacao!quantidade
                    Else
                        VlrIPI = TBCotacao!preco_unitario * TBCotacao!quantidade
                        VlrIPI = Format((VlrIPI * IntIPI) / 100, "###,##0.00")
                    End If
                    TBFIltro.Close
                    TBCotacao!dbl_valoripi = Format(VlrIPI, "###,##0.00")
                Else
                    TBCotacao!int_IPI = 0
                    TBCotacao!dbl_valoripi = 0
                End If
                
                If .txtuf <> "" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Empresa where Codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Carregar_CFOP_ST = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        ProcVerifCFOPST TBCotacao!ID_CF, .txtuf
                        If Valido = True Then
                            TBCotacao!ID_CFOP = IDAntigo
                            TBCotacao!txt_CST = Letra
                        End If
                    End If
                End If
                
                If TemICMS = "SIM" Then
                    TBCotacao!IntICMS = IntICMS
                    VlrIPI = TBCotacao!preco_unitario * TBCotacao!quantidade
                    ProcCalculaBC .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), IIf(IsNull(TBCotacao!ID_CFOP), 0, TBCotacao!ID_CFOP), 0, VlrIPI, IIf(IsNull(TBCotacao!dbl_valoripi), 0, TBCotacao!dbl_valoripi), SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBCotacao!txt_CST), "", TBCotacao!txt_CST), "P", 0, ""
                    TBCotacao!dbl_Valor_ICMS = Format((BC * IntICMS) / 100, "###,##0.00")
                Else
                    TBCotacao!IntICMS = 0
                    TBCotacao!dbl_Valor_ICMS = 0
                End If
            End If
            
            'Carrega comissão
            TBCotacao!Comissao = 0
            Set TBExecucao = CreateObject("adodb.recordset")
            TBExecucao.Open "select * from Vendas_Vendedores where N_Vendedor = " & .txtVE, Conexao, adOpenKeyset, adLockOptimistic
            If TBExecucao.EOF = False Then
                If TBExecucao!tipocomissao <> "" And IsNull(TBExecucao!tipocomissao) = False Then
                    If TBExecucao!tipocomissao = "V" Then
                        TBCotacao!Comissao = TBExecucao!Comissao
                    Else
                        Set TBCFOP = CreateObject("adodb.recordset")
                        If TBExecucao!tipocomissao = "C" Then TBCFOP.Open "select * from Vendas_Vendedores_Clientes where IDVendedor = " & TBExecucao!ID & " and IDCliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                        If TBExecucao!tipocomissao = "P" Then TBCFOP.Open "select * from Vendas_Vendedores_Produto where IDVendedor = " & TBExecucao!ID & " and IDProduto = " & TBItem!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
                        If TBExecucao!tipocomissao = "CP" Then TBCFOP.Open "select Vendas_Vendedores_Produto.comissao from Vendas_Vendedores_Produto INNER JOIN Vendas_Vendedores_Clientes on Vendas_Vendedores_Produto.idcliente = Vendas_Vendedores_Clientes.Id where Vendas_Vendedores_Produto.IDVendedor = " & TBExecucao!ID & " and Vendas_Vendedores_Produto.IDProduto = " & TBItem!Codproduto & " and Vendas_Vendedores_clientes.IDCliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                        If TBCFOP.EOF = False Then TBCotacao!Comissao = TBCFOP!Comissao
                        TBCFOP.Close
                    End If
                End If
            End If
            TBExecucao.Close
        End With
        TBCotacao!preco_unitario_desconto = IIf(IsNull(TBCotacao!preco_unitario), 0, TBCotacao!preco_unitario)
        TBCotacao!preco_lote = Format(IIf(IsNull(TBCotacao!preco_unitario), 0, TBCotacao!preco_unitario) * TBCotacao!quantidade, "###,##0.00")
        
        'Impostos
        Valor_total = TBCotacao!preco_lote
        Valor_IPI = IIf(IsNull(TBCotacao!dbl_valoripi), 0, TBCotacao!dbl_valoripi)
        
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            'Empresa
            ProcControleImposto IIf(IsNull(TBCotacao!ID_CFOP), 0, TBCotacao!ID_CFOP), IIf(.txtIDcliente = "", 0, .txtIDcliente)
            ProcVerifImpostosEmpresa .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), 0, "", False, 0, False, IIf(Vendas_PI = True, frmVendas_PI.TabelaSN_PI, frmVendas_proposta.TabelaSN_Proposta), 0
            If PI_Servicos = True Then
                'Novo cálculo simples nacional 2018
                TBCotacao!DAS = DAS
                If DAS <> 0 Then TBCotacao!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBCotacao!Total_DAS = 0
                TBCotacao!PIS_Serv = PIS_Serv
                If PIS_Serv <> 0 Then TBCotacao!Total_PIS_serv = Format((Valor_total * PIS_Serv) / 100, "###,##0.00") Else TBCotacao!Total_PIS_serv = 0
                TBCotacao!Cofins_Serv = Cofins_Serv
                If Cofins_Serv <> 0 Then TBCotacao!Total_Cofins_serv = Format((Valor_total * Cofins_Serv) / 100, "###,##0.00") Else TBCotacao!Total_Cofins_serv = 0
                TBCotacao!CSLL_Serv = CSLL_Serv
                If CSLL_Serv <> 0 Then TBCotacao!Total_CSLL_serv = Format((Valor_total * CSLL_Serv) / 100, "###,##0.00") Else TBCotacao!Total_CSLL_serv = 0
                TBCotacao!ISS = ISS_Serv
                If ISS_Serv <> 0 Then TBCotacao!VlrISS = Format((Valor_total * ISS_Serv) / 100, "###,##0.00") Else TBCotacao!VlrISS = 0
                TBCotacao!INSS_Serv = INSS_Serv
                If INSS_Serv <> 0 Then TBCotacao!Total_INSS_serv = Format((Valor_total * INSS_Serv) / 100, "###,##0.00") Else TBCotacao!Total_INSS_serv = 0
                TBCotacao!IRPJ_Serv = IRPJ_Serv
                If IRPJ_Serv <> 0 Then TBCotacao!Total_IRPJ_serv = Format((Valor_total * IRPJ_Serv) / 100, "###,##0.00") Else TBCotacao!Total_IRPJ_serv = 0
                TBCotacao!IRRF_Serv = IRRF_Serv
                If IRRF_Serv <> 0 Then TBCotacao!Total_IRRF_serv = Format((Valor_total * IRRF_Serv) / 100, "###,##0.00") Else TBCotacao!Total_IRRF_serv = 0
                TBCotacao!cpp = CPP_Serv
                If CPP_Serv <> 0 Then TBCotacao!Total_CPP = Format((Valor_total * CPP_Serv) / 100, "###,##0.00") Else TBCotacao!Total_CPP = 0
            Else
                'Novo cálculo simples nacional 2018
                TBCotacao!DAS = DAS
                If DAS <> 0 Then TBCotacao!Total_DAS = Format((Valor_total * DAS) / 100, "###,##0.00") Else TBCotacao!Total_DAS = 0
                TBCotacao!PIS_Prod = PIS_Prod
                If PIS_Prod <> 0 Then TBCotacao!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00") Else TBCotacao!Total_PIS_prod = 0
                TBCotacao!Cofins_Prod = Cofins_Prod
                If Cofins_Prod <> 0 Then TBCotacao!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00") Else TBCotacao!Total_Cofins_prod = 0
                TBCotacao!CSLL_Prod = CSLL_Prod
                If CSLL_Prod <> 0 Then TBCotacao!Total_CSLL_prod = Format((Valor_total * CSLL_Prod) / 100, "###,##0.00") Else TBCotacao!Total_CSLL_prod = 0
                TBCotacao!IRPJ_Prod = IRPJ_Prod
                If IRPJ_Prod <> 0 Then TBCotacao!Total_IRPJ_prod = Format((Valor_total * IRPJ_Prod) / 100, "###,##0.00") Else TBCotacao!Total_IRPJ_prod = 0
                TBCotacao!cpp = CPP_Prod
                If CPP_Prod <> 0 Then TBCotacao!Total_CPP = Format((Valor_total * CPP_Prod) / 100, "###,##0.00") Else TBCotacao!Total_CPP = 0
                
                TBCotacao!BC_ICMS = 0
                TBCotacao!BC_ICMS_ST = 0
                TBCotacao!Valor_ICMS_ST = 0
                If IsNull(TBCotacao!ID_CF) = False And TBCotacao!ID_CF <> "" Then
                    ProcValorImposto .txtCotacao, TBCotacao!ID_CF, IIf(.txtIDcliente = "", 0, .txtIDcliente), .txtCliente, .txtuf, .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), False, IIf(IsNull(TBCotacao!ID_CFOP), 0, TBCotacao!ID_CFOP), IIf(Vendas_PI = True, frmVendas_PI.RegimeEmpresa_PI, frmVendas_proposta.RegimeEmpresa_Proposta)
                    ProcControleImposto IIf(IsNull(TBCotacao!ID_CFOP), 0, TBCotacao!ID_CFOP), IIf(.txtIDcliente = "", 0, .txtIDcliente)
                    ProcCalculaBC .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), TBCotacao!ID_CFOP, 0, TBCotacao!preco_lote, TBCotacao!dbl_valoripi, SomarIPI, SomarIPIST, TemReducaoBC, False, IIf(IsNull(TBCotacao!txt_CST), "", TBCotacao!txt_CST), IIf(PI_Produtos = True, "P", "S"), 0, ""
                    If TemICMS = "SIM" And TBCotacao!dbl_Valor_ICMS <> 0 Then TBCotacao!BC_ICMS = BC
                    
                    If IsNull(TBCotacao!txt_CST) = False Then
                        ProcSubstituicaoTributaria .txtuf, TBCotacao!txt_CST, TBCotacao!ID_CF, IIf(.txtIDcliente = "", 0, .txtIDcliente), .txtCliente, TBCotacao!preco_unitario_desconto, TBCotacao!quantidade, BC, BCST, 0, 0, 0, False, False, 0
                        TBCotacao!Valor_ICMS_ST = ICMSCST
                        If ICMSCST <> 0 Then TBCotacao!BC_ICMS_ST = BCICMSCST
                    End If
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from tbl_classificacaofiscal where Idclass = " & TBCotacao!ID_CF, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        'Verifica se a CF tem retenção de PIS/Cofins, destaca PIS/Cofins e grava no produto
                        If TBFI!Retem_PIS_Cofins = True Then
                            TBCotacao!Valor_Retencao_PIS = Format((Valor_total * IIf(IsNull(TBFI!PIS), 0, TBFI!PIS)) / 100, "###,##0.00")
                            TBCotacao!Valor_Retencao_Cofins = Format((Valor_total * IIf(IsNull(TBFI!Cofins), 0, TBFI!Cofins)) / 100, "###,##0.00")
                        End If
                        
                        If Regime <> 1 Then
                            PIS_Prod = IIf(IsNull(TBFI!PIS_destaca), 0, TBFI!PIS_destaca)
                            Cofins_Prod = IIf(IsNull(TBFI!Cofins_destaca), 0, TBFI!Cofins_destaca)
                            If PIS_Prod <> 0 Then
                                TBCotacao!PIS_Prod = PIS_Prod
                                TBCotacao!Total_PIS_prod = Format((Valor_total * PIS_Prod) / 100, "###,##0.00")
                            End If
                            If Cofins_Prod <> 0 Then
                                TBCotacao!Cofins_Prod = Cofins_Prod
                                TBCotacao!Total_Cofins_prod = Format((Valor_total * Cofins_Prod) / 100, "###,##0.00")
                            End If
                        End If
                    End If
                    TBFI.Close
                End If
            End If
        End With
            
        'Calcula comissão
        Qtde = 0
        Qtd = 0
        If TBCotacao!Comissao <> 0 Then
            Qtde = TBCotacao!Comissao
            Qtd = TBCotacao!preco_lote
            Qtd = (Qtd * Qtde) / 100
            TBCotacao!ValorComissao = Qtd
        Else
            TBCotacao!ValorComissao = 0
        End If
    End If
    TBItem.Close
    TBCotacao.Update
    TBCotacao.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
'Verifica se o produto/serviço pertence ao cliente
If Vendas_Programacao = True Then
    With frmVendas_programacao
        IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        IDCliente = .txtID_cli
        Cliente = .txtCliente
    End With
Else
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        IDCliente = .txtIDcliente
        Cliente = .txtCliente
    End With
End If
If IDCliente <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & ListView1.SelectedItem & " and IDCliente <> 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Projproduto_clientes where Codproduto = " & ListView1.SelectedItem & " and IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            If PI_Servicos = True Or Proposta_Servicos = True Then NomeCampo = "serviço" Else NomeCampo = "Produto"
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from empresa where codigo = " & IDempresa & " and Bloquear_produtos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                If USMsgBox("Este " & NomeCampo & " não pertence ao cliente " & Cliente & ", deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                    TBCiclo.Close
                    Exit Sub
                End If
            Else
                USMsgBox ("Este " & NomeCampo & " não pertence ao cliente " & Cliente & "."), vbExclamation, "CAPRIND v5.0"
                TBCiclo.Close
                Exit Sub
            End If
            TBCiclo.Close
            
        End If
        TBProduto.Close
    End If
    TBAbrir.Close
End If

If Vendas_Programacao = True Then
    With frmVendas_programacao
        .txtCodigo = ListView1.SelectedItem.ListSubItems.Item(1)
        .ProcPuxaDadosProduto
    End With
Else
    If PI_Produtos = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            .ProcLimparProdutos True
            .txtNomenclatura = ListView1.SelectedItem.ListSubItems.Item(1)
            .ProcPuxaDadosProduto
        End With
    Else
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            .ProcLimparServicos True
            .txtcodservico = ListView1.SelectedItem.ListSubItems.Item(1)
            .ProcPuxadadosServico
        End With
    End If
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunCheckEditStatus() As Boolean
On Error GoTo tratar_erro
Dim hWnd As Long
Dim strClassName As String
Dim intPos As Integer

strClassName = Space(256)
hWnd = GetFocus()
Call GetClassName(hWnd, strClassName, 256)
intPos = InStr(1, strClassName, Chr(0))
strClassName = Left(strClassName, intPos - 1)
If strClassName = "ThunderRT6TextBox" Then FunCheckEditStatus = True    'Editing Else    FunCheckEditStatus = False

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
