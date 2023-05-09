VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRequisicao_materiais 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Requisição de materiais"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
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
      ItemData        =   "frmRequisicao_materiais.frx":0000
      Left            =   270
      List            =   "frmRequisicao_materiais.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1710
      Width           =   4080
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
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
      TabCaption(0)   =   "Requisição de materiais"
      TabPicture(0)   =   "frmRequisicao_materiais.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "txtid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(3)=   "USToolBar1"
      Tab(0).Control(4)=   "Lista_req"
      Tab(0).Control(5)=   "PBLista(0)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Lista de produtos"
      TabPicture(1)   =   "frmRequisicao_materiais.frx":0020
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Framelista"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lista"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtIDLista"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame1(28)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1(27)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Valores por centro de custo"
      TabPicture(2)   =   "frmRequisicao_materiais.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PBLista(2)"
      Tab(2).Control(1)=   "Lista2"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   585
         Index           =   27
         Left            =   12975
         TabIndex        =   75
         Top             =   9450
         Width           =   2325
         Begin VB.ComboBox Cmb_opcao_lista_Produtos 
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
            ItemData        =   "frmRequisicao_materiais.frx":0058
            Left            =   180
            List            =   "frmRequisicao_materiais.frx":0065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   210
            Width           =   1965
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Index           =   28
         Left            =   75
         TabIndex        =   73
         Top             =   9450
         Width           =   12900
         Begin DrawSuite2022.USProgressBar PBLista 
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   74
            Top             =   210
            Width           =   12525
            _ExtentX        =   22093
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   63
         Top             =   9090
         Width           =   15225
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
            ItemData        =   "frmRequisicao_materiais.frx":008E
            Left            =   6840
            List            =   "frmRequisicao_materiais.frx":0098
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   180
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
            Left            =   2880
            TabIndex        =   9
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
            TabIndex        =   10
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   14
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmRequisicao_materiais.frx":00B0
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
            TabIndex        =   13
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmRequisicao_materiais.frx":3854
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
            TabIndex        =   11
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
            TabIndex        =   12
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmRequisicao_materiais.frx":735D
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
            TabIndex        =   15
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmRequisicao_materiais.frx":B44C
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
            Index           =   3
            Left            =   3510
            TabIndex        =   86
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
            Left            =   5520
            TabIndex        =   79
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label17 
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   64
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtIDLista 
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
         Height          =   330
         Left            =   1800
         TabIndex        =   62
         Text            =   "0"
         ToolTipText     =   "IDLista."
         Top             =   5280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtid 
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
         Left            =   -72480
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5130
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1965
         Index           =   0
         Left            =   -74925
         TabIndex        =   41
         Top             =   1320
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
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1665
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
            Left            =   12480
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   2535
         End
         Begin VB.TextBox txtrequisicao 
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
            Left            =   4290
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número da requisição."
            Top             =   390
            Width           =   1305
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
            Height          =   315
            Left            =   5610
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1785
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
            Left            =   7410
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   795
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
            Left            =   8220
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2565
         End
         Begin VB.TextBox txtobs 
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
            Height          =   825
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            ToolTipText     =   "Observações da requisição."
            Top             =   1020
            Width           =   14835
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Index           =   2
            Left            =   12757
            TabIndex        =   77
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   10908
            TabIndex        =   76
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   1890
            TabIndex        =   49
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº requisição"
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
            Left            =   4387
            TabIndex        =   48
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   6237
            TabIndex        =   46
            Top             =   180
            Width           =   555
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
            Left            =   9066
            TabIndex        =   45
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
            Left            =   7659
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   7155
            TabIndex        =   42
            Top             =   810
            Width           =   870
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   67
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
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
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar (F7)"
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
         ButtonWidth8    =   44
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Validação"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Validar/Cancelar validação (F8)"
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
         ButtonLeft9     =   399
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   454
         ButtonTop10     =   2
         ButtonWidth10   =   59
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
         ButtonLeft11    =   515
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
         ButtonLeft12    =   519
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft13    =   557
         ButtonTop13     =   2
         ButtonWidth13   =   30
         ButtonHeight13  =   21
         ButtonUseMaskColor13=   0   'False
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonKey14     =   "14"
         ButtonAlignment14=   2
         BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState14   =   5
         ButtonLeft14    =   589
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         ButtonUseMaskColor14=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   13290
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmRequisicao_materiais.frx":ECD8
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista_req 
         Height          =   5775
         Left            =   -74925
         TabIndex        =   8
         Top             =   3300
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   10186
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
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº requisição"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   5204
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Validado"
            Object.Width           =   1499
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   75
         TabIndex        =   68
         Top             =   360
         Width           =   15225
         _ExtentX        =   26855
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
         ButtonKey2      =   "3"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   86
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   127
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
         ButtonLeft5     =   189
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
         ButtonLeft6     =   246
         ButtonTop6      =   2
         ButtonWidth6    =   55
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Centro de custo"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Centro de custo (F7)"
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
         ButtonLeft7     =   303
         ButtonTop7      =   2
         ButtonWidth7    =   97
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Autorizar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Autorizar (F8)"
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
         ButtonLeft8     =   402
         ButtonTop8      =   2
         ButtonWidth8    =   61
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
         ButtonLeft9     =   465
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
         ButtonLeft10    =   469
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   507
         ButtonTop11     =   2
         ButtonWidth11   =   30
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
         ButtonLeft12    =   539
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13260
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmRequisicao_materiais.frx":171AE
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   5205
         Left            =   75
         TabIndex        =   38
         Top             =   4230
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   9181
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   9
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
            Object.Width           =   7735
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Vlr. unitário"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Vlr. total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView Lista2 
         Height          =   9405
         Left            =   -74925
         TabIndex        =   82
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   16589
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "T"
            Text            =   "Centro de custo"
            Object.Width           =   24051
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "V"
            Text            =   "Valor total"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Frame Framelista 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2895
         Left            =   75
         TabIndex        =   50
         Top             =   1320
         Width           =   15225
         Begin VB.TextBox Txt_ordem 
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
            Left            =   13860
            TabIndex        =   28
            ToolTipText     =   "Número da ordem de produção."
            Top             =   1020
            Width           =   1185
         End
         Begin VB.TextBox txtvlrTotal 
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
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Valor total."
            Top             =   1620
            Width           =   1335
         End
         Begin VB.TextBox txtvlrUnit 
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
            Height          =   315
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Valor unitário."
            Top             =   1620
            Width           =   1245
         End
         Begin VB.ComboBox Cmb_un_com 
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
            ItemData        =   "frmRequisicao_materiais.frx":1DDE7
            Left            =   12990
            List            =   "frmRequisicao_materiais.frx":1DDE9
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Unidade comercial."
            Top             =   1020
            Width           =   855
         End
         Begin VB.CommandButton cmdfiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   7770
            Picture         =   "frmRequisicao_materiais.frx":1DDEB
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Filtrar por código interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmdescolheraut 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14730
            Picture         =   "frmRequisicao_materiais.frx":1E206
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Selecionar autorizado."
            Top             =   1620
            Width           =   315
         End
         Begin VB.TextBox txtAutorizado 
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
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Autorizado por."
            Top             =   1620
            Width           =   3900
         End
         Begin VB.TextBox txtData_Autorizacao 
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
            Left            =   9750
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Data da autorização."
            Top             =   1620
            Width           =   1035
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Criar novo produto/item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   11700
            TabIndex        =   69
            Top             =   180
            Width           =   3345
            Begin VB.CheckBox chkManual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cód. manual ?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   120
               TabIndex        =   22
               Top             =   270
               Width           =   1335
            End
            Begin VB.CheckBox chkAuto 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cód. automático ?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1620
               TabIndex        =   23
               Top             =   270
               Width           =   1605
            End
         End
         Begin VB.TextBox txtResponsavel_item 
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
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4665
         End
         Begin VB.TextBox txtData_item 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1035
         End
         Begin VB.TextBox txtstatus_item 
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   3075
         End
         Begin VB.CommandButton cmdEscolher_item 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8100
            Picture         =   "frmRequisicao_materiais.frx":1E308
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Localizar produtos/itens."
            Top             =   390
            Width           =   315
         End
         Begin VB.ComboBox cmbun 
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
            ItemData        =   "frmRequisicao_materiais.frx":1E40A
            Left            =   12120
            List            =   "frmRequisicao_materiais.frx":1E40C
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Unidade de estoque."
            Top             =   1020
            Width           =   855
         End
         Begin VB.TextBox txtQE 
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade em estoque."
            Top             =   1620
            Width           =   1005
         End
         Begin VB.TextBox txtQR 
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
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            ToolTipText     =   "Quantidade requisitada."
            Top             =   1620
            Width           =   1005
         End
         Begin VB.TextBox txtdesenho 
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
            Left            =   5910
            TabIndex        =   18
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1875
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
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1020
            Width           =   7005
         End
         Begin VB.TextBox txtObs_item 
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
            Height          =   555
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            ToolTipText     =   "Observações."
            Top             =   2205
            Width           =   14865
         End
         Begin VB.ComboBox cmbfamilia 
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
            Left            =   7200
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   1020
            Width           =   4905
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
            Left            =   4830
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            ToolTipText     =   "Centro de custo."
            Top             =   1620
            Width           =   4905
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "OP"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   14347
            TabIndex        =   84
            Top             =   810
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vlr.total"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3862
            TabIndex        =   81
            Top             =   1410
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vlr. unitario"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2445
            TabIndex        =   80
            Top             =   1410
            Width           =   825
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Un com.*"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   72
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label15 
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
            Index           =   1
            Left            =   10095
            TabIndex        =   71
            Top             =   1410
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Autorizado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12360
            TabIndex        =   70
            Top             =   1410
            Width           =   780
         End
         Begin VB.Label Label9 
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
            Left            =   525
            TabIndex        =   61
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label6 
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
            Left            =   3105
            TabIndex        =   60
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Codigo interno*"
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
            Index           =   3
            Left            =   6180
            TabIndex        =   59
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Qtde.*"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1447
            TabIndex        =   58
            Top             =   1410
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   3292
            TabIndex        =   57
            Top             =   810
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Família*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9367
            TabIndex        =   56
            Top             =   810
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Un est.*"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   12240
            TabIndex        =   55
            Top             =   810
            Width           =   615
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. est."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   307
            TabIndex        =   54
            Top             =   1410
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Observação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7177
            TabIndex        =   53
            Top             =   2010
            Width           =   870
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   9780
            TabIndex        =   52
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
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
            Left            =   6705
            TabIndex        =   51
            Top             =   1410
            Width           =   1155
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Index           =   2
         Left            =   -74925
         TabIndex        =   83
         Top             =   9750
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
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Index           =   0
         Left            =   -74925
         TabIndex        =   85
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
   End
End
Attribute VB_Name = "frmRequisicao_materiais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_requisicao_material  As Boolean 'OK
Dim Novo_requisicao_material1 As Boolean 'OK
Public StrSql_Localizar_Requisicao As String 'OK
Dim TBLISTA_Requisicao     As ADODB.Recordset 'OK
Dim IDRequisicao As Integer
Dim status As String

Private Sub ProcVerificaStatusRM()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Requisicao_materiais_lista where  idrequisicao = " & IDRequisicao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
  status = "ABERTA"
Exit Sub
End If
TBAbrir.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Requisicao_materiais_lista where  Status <> 'RETIRADO' and idrequisicao = " & IDRequisicao, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
status = "RETIRADA"
Else
status = "ABERTA"
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkAuto_Click()
On Error GoTo tratar_erro

If chkAuto.Value = 1 Then
    chkManual.Value = 0
    Procliberacampos
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkManual_Click()
On Error GoTo tratar_erro

If chkManual.Value = 1 Then
    chkAuto.Value = 0
    Procliberacampos
Else
    ProcBloqueiaCampos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Procliberacampos()
On Error GoTo tratar_erro

With txtdescricao
    .Locked = False
    .TabStop = True
End With
With cmbfamilia
    .Locked = False
    .TabStop = True
End With
With cmbun
    .Locked = False
    .TabStop = True
End With
With Cmb_un_com
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With txtdescricao
    .Locked = True
    .TabStop = False
End With
With cmbfamilia
    .Locked = True
    .TabStop = False
End With
With cmbun
    .Locked = True
    .TabStop = False
End With
With Cmb_un_com
    .Locked = True
    .TabStop = False
End With
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

If txtrequisicao = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Requisicao_materiais order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Requisicao = '" & txtrequisicao & "'")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        ProcLimparTudo
        txtrequisicao = TBLISTA!requisicao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Requisicao_materiais where Requisicao = '" & txtrequisicao & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcPuxaDados
        ProcCarregaLista_Item
    Else
        USMsgBox ("Fim dos cadastros de requisição."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_requisicao_material = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtrequisicao = "" Then
    Acao = "copiar"
    NomeCampo = "a requisição"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_requisicao_material = True Then
    USMsgBox ("Salve a requisição antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar esta requisição?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    IDpedido = txtId
    Set TBSolicitacao = CreateObject("adodb.recordset")
    TBSolicitacao.Open "Select * from Requisicao_materiais", Conexao, adOpenKeyset, adLockOptimistic
    TBSolicitacao.AddNew
    TBSolicitacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBSolicitacao!Responsavel = pubUsuario
    TBSolicitacao!Data = Format(Date, "dd/mm/yy")
    TBSolicitacao!Obs = txtObs
    TBSolicitacao!status = "ABERTA"
    ProcCriarNovoNumero
    Novo_requisicao_material = True
    txtrequisicao = a
    TBSolicitacao!requisicao = txtrequisicao
    TBSolicitacao.Update
    txtId = TBSolicitacao!ID
    TBSolicitacao.Close
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Requisicao_materiais_lista where IDRequisicao = " & IDpedido & " order by IdLista", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Requisicao_materiais_lista", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!Data = Date
            TBGravar!Responsavel = pubUsuario
            TBGravar!Desenho = TBAbrir!Desenho
            TBGravar!Descricao = TBAbrir!Descricao
            TBGravar!status = "REQUISIT."
            TBGravar!Un = TBAbrir!Un
            TBGravar!Unidade_com = TBAbrir!Unidade_com
            TBGravar!IDRequisicao = txtId
            TBGravar!Quant = TBAbrir!Quant
            TBGravar!Familia = TBAbrir!Familia
            TBGravar!Obs = TBAbrir!Obs
            TBGravar!ID_CC = TBAbrir!ID_CC
            TBGravar.Update
            TBGravar.Close
            TBAbrir.MoveNext
        Loop
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Requisicao_materiais where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcLimpaCampos
        ProcPuxaDados
    End If
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista_req.ListItems.Count <> 0 Then
        Lista_req.SelectedItem = Lista_req.ListItems(CodigoLista)
        Lista_req.SetFocus
    End If
    USMsgBox ("Requisição copiada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Requisição de materiais"
    Evento = "Novo"
    ID_documento = txtId
    Documento = "Nº requisição: " & txtrequisicao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Frame1(0).Enabled = True
    Novo_requisicao_material = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lista_req
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

Private Sub Cmb_opcao_lista_Produtos_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar2
    Select Case Cmb_opcao_lista_Produtos
        Case "Excluir"
            .ButtonState(3) = 0
            .ButtonState(7) = 5
            .ButtonState(8) = 5
        Case "Centro de custo"
            .ButtonState(3) = 5
            .ButtonState(7) = 0
            .ButtonState(8) = 5
        Case "Autorizar"
            .ButtonState(3) = 5
            .ButtonState(7) = 5
            .ButtonState(8) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEscolher_item_Click()
On Error GoTo tratar_erro

frmRequisicao_materiais_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoProduto()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus <> "Aberta" Then
    USMsgBox ("Não é permitido criar novo produto, pois a requisição está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "requisição", "produto", False) = False Then Exit Sub

ProcLimpaCampos_produto
Novo_requisicao_material1 = True
Framelista.Enabled = True
txtdesenho.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtrequisicao = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Requisicao_materiais order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Requisicao = '" & txtrequisicao & "'")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        ProcLimparTudo
        txtrequisicao = TBLISTA!requisicao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Requisicao_materiais where Requisicao = '" & txtrequisicao & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcPuxaDados
        ProcCarregaLista_Item
    Else
        USMsgBox ("Fim dos cadastros de requisição."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_requisicao_material = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdescolheraut_Click()
On Error GoTo tratar_erro

If FunVerificaRegistroValidado("Requisicao_materiais", "ID = " & txtId, "requisição", "do produto", "autorizar/cancelar autorização", False, False) = False Then Exit Sub

If txtStatus = "Retirada" Then
    USMsgBox ("Não é permitido autorizar/cancelar autorização, pois a requisição está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtstatus_item <> "Requisitado" Then
    USMsgBox ("Não é permitido autorizar/cancelar autorização, pois o produto está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_requisicao_material1 = True Then
    USMsgBox ("Salve o produto antes de autorizar/cancelar autorização."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID_CC from Requisicao_materiais_lista where IdLista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!ID_CC) = True Or TBAbrir!ID_CC = "0" Or TBAbrir!ID_CC = "" Then
        USMsgBox ("Não é permitido autorizar/cancelar autorização, pois este produto não possui centro de custo."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
End If
TBAbrir.Close
Compras_Cotacao = False
Estoque_Requisicao = True
frmCompras_Requisicao_aut.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

ProcCarregaProduto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaProduto()
On Error GoTo tratar_erro

If txtdesenho <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select E.*, P.Unidade_com from Estoque_produtos E INNER JOIN Projproduto P ON P.Desenho = E.Desenho where E.desenho = '" & txtdesenho.Text & "' and E.Liberado = 'SIM' and (LEFT(E.status, 7) = 'ENTRADA' or E.status = 'CONSIGNAÇÃO RECEBIDA')", Conexao, adOpenKeyset, adLockOptimistic
    StrSql = "Select E.*, P.Unidade_com from Estoque_produtos E INNER JOIN Projproduto P ON P.Desenho = E.Desenho where E.desenho = '" & txtdesenho.Text & "' and E.Liberado = 'SIM' and (LEFT(E.status, 7) = 'ENTRADA' or E.status = 'CONSIGNAÇÃO RECEBIDA')"
    'Debug.print StrSql
    
    
    If TBProduto.EOF = False Then
        txtdesenho = TBProduto!Desenho
        txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        NomeCampo = "a família"
        If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia.Text = TBProduto!Classe
        NomeCampo = "a unidade de estoque"
        If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun.Text = TBProduto!Unidade
        NomeCampo = "a unidade comercial"
        If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com.Text = TBProduto!Unidade_com
2:
        txtQE = Format(FunVerificaQtdeEstoque(txtdesenho, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
        
        ProcBloqueiaCampos
        With cmbun
            .Locked = True
            .TabStop = False
        End With
        With Cmb_un_com
            .Locked = True
            .TabStop = False
        End With
        procCarrega_CentroDeCusto
    Else
        txtdescricao.Text = ""
        cmbun.ListIndex = -1
        Cmb_un_com.ListIndex = -1
        cmbfamilia.ListIndex = -1
        txtQE.Text = "0,0000"
        Procliberacampos
    End If
Else
    Procliberacampos
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " desse registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifValorTotalSaida(IDlista As Long)
On Error GoTo tratar_erro

valor = 0
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select Sum(VlrTotal) as Valor from Estoque_movimentacao where ID_prod_RM = " & IDlista & " and Documento = '" & txtrequisicao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    valor = IIf(IsNull(TBEstoque!valor), 0, TBEstoque!valor)
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Requisicao.AbsolutePage <> 2 Then
    If TBLISTA_Requisicao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Requisicao.PageCount - 1)
    Else
        TBLISTA_Requisicao.AbsolutePage = TBLISTA_Requisicao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Requisicao.AbsolutePage)
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
    TBLISTA_Requisicao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Requisicao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Requisicao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Requisicao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Requisicao.AbsolutePage <> -3 Then
    If TBLISTA_Requisicao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Requisicao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Requisicao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Requisicao.AbsolutePage = TBLISTA_Requisicao.PageCount
ProcExibePagina (TBLISTA_Requisicao.AbsolutePage)

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
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiar
            Case vbKeyF8: If Cmb_opcao_lista = "Validação" Then ProcExcluir
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoProduto
            Case vbKeyF3: ProcSalvarProduto
            Case vbKeyF4: If Cmb_opcao_lista_Produtos = "Excluir" Then ProcExcluirProduto
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista_Produtos = "Centro de custo" Then procCentroDeCusto
            Case vbKeyF8: If Cmb_opcao_lista_Produtos = "Autorizar" Then procAutorizar
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            'Case vbKeyF1: ProcAjuda
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

If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
txtId = TBAbrir!ID
Caption = "Estoque - Requisição de materiais - (Requisição : " & IIf(IsNull(TBAbrir!requisicao), "", TBAbrir!requisicao) & ")"
txtrequisicao = IIf(IsNull(TBAbrir!requisicao), "", TBAbrir!requisicao)
txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel.Text = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
With txtStatus
    Select Case TBAbrir!status
        Case "ABERTA": .Text = "Aberta"
        Case "RETIRADA": .Text = "Retirada"
        Case "PARCIAL": .Text = "Retirada parcial"
    End Select
End With
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
Novo_requisicao_material = False
Frame1(0).Enabled = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaLista_Item()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If CC_RM = True Then TextoFiltro = " and ID_CC = " & frmRelatorios_Custos_Prev_Real.ID_CC Else TextoFiltro = ""
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Requisicao_materiais_lista Where IDRequisicao = " & txtId & TextoFiltro & " order by idlista desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(1).Min = 0
    PBLista(1).Max = TBLISTA.RecordCount
    PBLista(1).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Familia), "", TBLISTA!Familia)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Quant), "", Format(TBLISTA!Quant, "###,##0.0000"))
            
            ProcVerifValorTotalSaida TBLISTA!IDlista
            .Item(.Count).SubItems(6) = Format(valor / TBLISTA!Quant, "###,##0.0000000000")
            .Item(.Count).SubItems(7) = Format(valor, "###,##0.00")
            
            If IsNull(TBLISTA!status) = False Then
                Select Case TBLISTA!status
                    Case "REQUISIT.": .Item(.Count).SubItems(8) = "Requisitado"
                    Case "RETIRADO": .Item(.Count).SubItems(8) = "Retirado"
                    Case "PARCIAL": .Item(.Count).SubItems(8) = "Retirado parcial"
                    Case "CANCELADO": .Item(.Count).SubItems(8) = "Cancelado"
                End Select
            End If
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(1).Value = Contador
    Loop
End If
TBLISTA.Close



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_produto()
On Error GoTo tratar_erro

TXTIDLista = 0
txtData_item = Format(Date, "dd/mm/yy")
txtResponsavel_item = pubUsuario
txtdesenho.Text = ""
txtstatus_item = "Requisitado"
chkAuto.Value = 0
chkManual.Value = 0
txtdescricao.Text = ""
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
Txt_ordem = ""
txtQE.Text = ""
txtQR.Text = ""
txtvlrUnit = ""
txtvlrTotal = ""
txtData_Autorizacao = ""
txtAutorizado = ""
txtObs_item = ""
CodigoLista1 = 0
procCarrega_CentroDeCusto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtrequisicao.Text = ""
txtData.Text = Format(Date, "dd/mm/yy")
txtResponsavel.Text = pubUsuario
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtStatus.Text = "Aberta"
txtObs.Text = ""
CodigoLista = 0
Caption = "Estoque - Requisição de materiais"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Procenviadadoslista()
On Error GoTo tratar_erro

TBCompras!IDRequisicao = txtId
TBCompras!Data = IIf(txtData_item = "", Date, txtData_item)
TBCompras!Responsavel = IIf(txtResponsavel_item = "", pubUsuario, txtResponsavel_item)
If txtstatus_item = "Requisitado" Then TBCompras!status = "REQUISIT."
TBCompras!Desenho = txtdesenho.Text
TBCompras!Quant = txtQR.Text
TBCompras!Familia = cmbfamilia.Text
TBCompras!Descricao = txtdescricao
TBCompras!Un = cmbun.Text
TBCompras!Unidade_com = Cmb_un_com.Text
TBCompras!Ordem = IIf(Txt_ordem = "", Null, Txt_ordem)
If Cmb_centro = "" Then
    TBCompras!ID_CC = Null
    TBCompras!Data_autorizacao = Null
    TBCompras!Autorizado = ""
Else
    TBCompras!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex)
End If
TBCompras!Obs = txtObs_item

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select ID_PC from projproduto where desenho = '" & txtdesenho & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBCompras!ID_PC = TBProduto!ID_PC
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarProduto()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtstatus_item <> "Requisitado" Then
    USMsgBox ("Não é permitido alterar este produto, pois o mesmo está " & txtstatus_item & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "requisição", "do produto", False) = False Then Exit Sub

If Framelista.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If chkAuto.Value = 0 And txtdesenho = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtdesenho.SetFocus
    Exit Sub
End If
If txtdescricao.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
If cmbfamilia.Text = "" Then
    NomeCampo = "a familia"
    ProcVerificaAcao
    cmbfamilia.SetFocus
    Exit Sub
End If
If cmbun.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    cmbun.SetFocus
    Exit Sub
End If
If Cmb_un_com.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
If Txt_ordem <> "" Then
    If FunVerifOP(Txt_ordem, Novo_requisicao_material1) = False Then
        Txt_ordem.SetFocus
        Exit Sub
    End If
End If
Qtde = IIf(txtQR = "", 0, txtQR)
If Qtde = 0 Then
    NomeCampo = "a quantidade requisitada"
    ProcVerificaAcao
    txtQR.SetFocus
    Exit Sub
End If
If chkAuto.Value = 1 Then ProcNovoProdutoAuto
If chkManual.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtN_Estoque.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManual
End If
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtdesenho.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = True Then
    USMsgBox ("Não foi encontrado nenhum produto com este código interno, favor marcar a opção de Criar novo item (cód. automático) antes de salvar."), vbExclamation, "CAPRIND v5.0"
    chkAuto.SetFocus
    TBProduto.Close
    Exit Sub
End If
TBProduto.Close
    
'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
Permitido = False
ID_CC = 0
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If Novo_requisicao_material1 = False Then
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select ID_CC from Requisicao_materiais_lista WHERE IDLista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then ID_CC = IIf(IsNull(TBCompras!ID_CC), 0, TBCompras!ID_CC)
        TBCompras.Close
    End If
            
    If Cmb_centro <> "" Then
        If ID_CC <> Cmb_centro.ItemData(Cmb_centro.ListIndex) Then
            Formulario = "Estoque/Autorização de centro de custo sem previsão"
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select ID_PC from projproduto where desenho = '" & txtdesenho & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = True Then
                If USMsgBox("O produto não possui conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
            Else
                Set TBCQ = CreateObject("adodb.recordset")
                TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & Cmb_centro.ItemData(Cmb_centro.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                If TBCQ.EOF = True Then
                    If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                Else
                    Permitido = True
                End If
                TBCQ.Close
            End If
            TBProduto.Close
            If Permitido = False Then Exit Sub
        End If
    End If
End If
TBTempo.Close

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Requisicao_materiais_lista WHERE IDLista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then TBCompras.AddNew
Procenviadadoslista
TBCompras.Update
TXTIDLista = TBCompras!IDlista
TBCompras.Close
If Novo_requisicao_material1 = True Then
    USMsgBox ("Novo produto cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto"
    ProcCarregaLista_Item
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto"
    ProcCarregaLista_Item
    If CodigoLista1 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista1)
        Lista.SetFocus
    End If
End If
Novo_requisicao_material1 = False
'==================================
Modulo = "Estoque/Requisição de materiais"
ID_documento = TXTIDLista
Documento = "Nº requisição: " & txtrequisicao
Documento1 = "Cód. interno: " & txtdesenho
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoAuto()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtdesenho = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtdesenho, "", 0, txtdescricao, txtdescricao, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtdesenho = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtdesenho, "", 0, txtdescricao, txtdescricao, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoManual()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtdesenho = FunCriaNovoProdServ(True, "", txtdesenho, "", 0, txtdescricao, txtdescricao, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtdesenho = FunCriaNovoProdServ(True, "", txtdesenho, "", 0, txtdescricao, txtdescricao, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluirProduto()
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
                If USMsgBox("Deseja realmente excluir este(s) produto(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Requisicao_materiais_lista where IDlista = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Estoque/Requisição de materiais"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº requisição: " & txtrequisicao
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_produto
    ProcCarregaLista_Item
    Framelista.Enabled = False
    Novo_requisicao_material1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 13, True
ProcCarregaToolBar2 Me, 15195, 12, True

Formulario = "Estoque/Requisição de materiais"
Direitos
SSTab1.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaFamiliaUN
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Validação"
Cmb_opcao_lista_Produtos = "Excluir"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Estoque/Requisição de materiais"
Direitos
ProcCarregaFamiliaUN
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamiliaUN()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
If TXTIDLista <> 0 Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from Requisicao_materiais_lista where idlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    If IsNull(TBFamilia!Familia) = False And TBFamilia!Familia <> "" Then cmbfamilia = TBFamilia!Familia
    If IsNull(TBFamilia!Un) = False And TBFamilia!Un <> "" Then cmbun = TBFamilia!Un
    If IsNull(TBFamilia!Unidade_com) = False And TBFamilia!Unidade_com <> "" Then Cmb_un_com = TBFamilia!Unidade_com
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362R" Then frmRequisicao_materiais_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro
   
If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmRequisicao_materiais_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza número das requisições
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from Requisicao_materiais order by ID", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                PBLista(0).Min = 0
                PBLista(0).Max = TBCompras.RecordCount
                PBLista(0).Value = 1
                Contador = 0
                Do While TBCompras.EOF = False
                    Ano = Right(Year(TBCompras!Data), 2)
                    If Right(TBCompras!requisicao, 3) <> "/" & Ano Then
                        IDAntigo = Right(TBCompras!requisicao, 5)
                        Conexao.Execute "Update Requisicao_materiais_lista Set IDrequisicao = " & TBCompras!ID & " where IDrequisicao = " & IDAntigo
                        
                        RequisicaoNovo = TBCompras!requisicao & "/" & Ano
                        TBCompras!requisicao = RequisicaoNovo
                    End If
                    TBCompras.Update
                    TBCompras.MoveNext
                    Contador = Contador + 1
                    PBLista(0).Value = Contador
                Loop
            End If
            TBCompras.Close
            
            Lista_req.ListItems.Clear
            ProcCarregaLista (1)
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Outros/Solicitação"
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

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmRequisicao_materiais_abrir.Show 1

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
With Lista_req
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) requisição(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Requisicao_materiais where id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Requisicao_materiais_lista where idrequisicao = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Estoque/Requisição de materiais"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº requisição: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) requisição(ões) antes de excluir ou alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Requisição(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcLimparTudo
    ProcCarregaLista (1)
    Novo_requisicao_material = False
    Frame1(0).Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = 0 Then
    USMsgBox ("Informe a requisição de materiais antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "Estoque_requisicao_materiais.rpt"
ProcImprimirRel "{Requisicao_materiais.id} = " & txtId, ""

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
Novo_requisicao_material = True
Frame1(0).Enabled = True
txtObs.SetFocus
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Framelista.Enabled = False
ProcLimpaCampos_produto
Lista.ListItems.Clear
Novo_requisicao_material1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarNovoNumero()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select requisicao from Requisicao_materiais where Year(data) = '" & Year(Date) & "' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Numero = Left(TBAbrir!requisicao, Len(TBAbrir!requisicao) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close
Ano = Right(Year(Date), 2)
a = "RM-" & FunTamanhoTextoZeroEsq(Numero, 5) & "/" & Ano

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_requisicao_material = True Then
    If USMsgBox("A requisição ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_requisicao_material = True Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End If
If Novo_requisicao_material1 = True Then
    If USMsgBox("O produto ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarProduto
        If Novo_requisicao_material1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_requisicao_material = False
Novo_requisicao_material1 = False
CC_RM = False
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
If Frame1(0).Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtStatus <> "Aberta" Then
    USMsgBox ("Não é permitido alterar esta requisição, pois a mesma está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Requisicao_materiais where ID = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then
    ProcCriarNovoNumero
    txtrequisicao = a
    TBCompras.AddNew
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesmo", "requisição de materiais", True) = False Then Exit Sub
End If
TBCompras!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBCompras!Responsavel = pubUsuario
TBCompras!Data = Date
TBCompras!status = "ABERTA"
TBCompras!requisicao = txtrequisicao
TBCompras!Obs = IIf(txtObs.Text = "", Null, txtObs.Text)
TBCompras.Update
txtId = TBCompras!ID
Caption = "Estoque - Requisição de materiais - (Requisição : " & IIf(IsNull(TBCompras!requisicao), "", TBCompras!requisicao) & ")"
TBCompras.Close
If Novo_requisicao_material = True Then
    USMsgBox ("Nova requisição cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    CamposFiltro = "RM.ID, RM.requisicao, RM.Data, RM.Responsavel, RM.Status, RM.DtValidacao, E.Empresa"
    StrSql_Localizar_Requisicao = "Select " & CamposFiltro & " from Requisicao_materiais RM INNER JOIN Empresa E ON E.Codigo = RM.ID_empresa where ID = " & txtId & " group by " & CamposFiltro
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista_req.ListItems.Count <> 0 Then
        Lista_req.SelectedItem = Lista_req.ListItems(CodigoLista)
        Lista_req.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Estoque/Requisição de materiais"
    ID_documento = txtId
    Documento = "Nº requisição: " & txtrequisicao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_requisicao_material = False

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
                If Cmb_opcao_lista_Produtos = "Autorizar" Then
                    If FunVerificaRegistroValidadoSemMsg("Requisicao_materiais", "ID = " & txtId, False) = False Then
                        GoTo Proximo
                    End If
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select ID_CC from Requisicao_materiais_lista where IdLista = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        If IsNull(TBAbrir!ID_CC) = True Or TBAbrir!ID_CC = "0" Or TBAbrir!ID_CC = "" Then
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                    End If
                    TBAbrir.Close
                End If
                
                If txtStatus = "Retirada" Then GoTo Proximo
                If .ListItems.Item(InitFor).SubItems(8) <> "Requisitado" Then GoTo Proximo
                If Cmb_opcao_lista_Produtos <> "Autorizar" Then
                    If FunVerificaRegistroValidadoSemMsg("Requisicao_materiais", "ID = " & txtId, True) = False Then GoTo Proximo
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
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista_Produtos = "Excluir" Then
                Texto = "excluir"
            ElseIf Cmb_opcao_lista_Produtos = "Autorizar" Then
                    If FunVerificaRegistroValidado("Requisicao_materiais", "ID = " & txtId, "requisição", "do produto", "autorizar/cancelar autorização", False, False) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select ID_CC from Requisicao_materiais_lista where IdLista = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        If IsNull(TBAbrir!ID_CC) = True Or TBAbrir!ID_CC = "0" Or TBAbrir!ID_CC = "" Then
                            USMsgBox ("Não é permitido autorizar/cancelar autorização, pois este produto não possui centro de custo."), vbExclamation, "CAPRIND v5.0"
                            TBAbrir.Close
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    End If
                    TBAbrir.Close
                    Texto = "autorizar/cancelar autorização"
                Else
                    Texto = "alterar centro de custo"
            End If
            
            If txtStatus = "Retirada" Then
                USMsgBox ("Não é permitido " & Texto & ", pois a requisição está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If .ListItems.Item(InitFor).SubItems(8) <> "Requisitado" Then
                USMsgBox ("Não é permitido " & Texto & IIf(Cmb_opcao_lista_Produtos = "Excluir", "este", "deste") & " produto, pois o status do mesmo está " & .ListItems.Item(InitFor).SubItems(6) & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If Cmb_opcao_lista_Produtos <> "Autorizar" Then
                If FunVerificaRegistroValidado("Requisicao_materiais", "ID = " & txtId, "requisição", "produto", Texto & IIf(Cmb_opcao_lista_Produtos = "Excluir", " este", " deste"), False, True) = False Then .ListItems.Item(InitFor).Checked = False
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

ProcLimpaCampos_produto
TXTIDLista = Lista.SelectedItem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Requisicao_materiais_lista where idlista = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    procCarregaDados_Itens
    CodigoLista1 = Lista.SelectedItem.index
End If
TBAbrir.Close
Novo_requisicao_material1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_req_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_req
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).SubItems(5) <> "Aberta" Then GoTo Proximo
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("Requisicao_materiais", "ID = " & .ListItems(InitFor), True) = False Then GoTo Proximo
                End If
                
                If Cmb_opcao_lista = "Validação" Then
                    If .ListItems.Item(InitFor).SubItems(6) = "Sim" Then
                        'Verifica se tem produto com CC aprovado
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select IdLista from Requisicao_materiais_lista where IDrequisicao = " & .ListItems(InitFor) & " and Data_autorizacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                        TBAbrir.Close
                    Else
                        'Verifica se a empresa exige centro de custo
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and CC_obrigatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select IDrequisicao from Requisicao_materiais_lista where IDrequisicao = " & .ListItems(InitFor) & " and ID_CC IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                TBAbrir.Close
                                TBFI.Close
                                GoTo Proximo
                            End If
                            TBAbrir.Close
                        End If
                        TBFI.Close
                    End If
                End If
                
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_req, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_req_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_req
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If .ListItems.Item(InitFor).SubItems(5) <> "Aberta" Then
                USMsgBox ("Não é permitido " & IIf(Cmb_opcao_lista = "Excluir", "excluir esta", "validar/cancelar validação desta") & " requisição, pois o status da mesma está " & .ListItems.Item(InitFor).SubItems(5) & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("Requisicao_materiais", "ID = " & .ListItems.Item(InitFor), "mesma", "requisição de material", "excluir esta ", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            
            If Cmb_opcao_lista = "Validação" Then
                If .ListItems.Item(InitFor).SubItems(6) = "Sim" Then
                    'Verifica se tem produto com CC aprovado
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select IdLista from Requisicao_materiais_lista where IDrequisicao = " & .ListItems(InitFor) & " and Data_autorizacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido cancelar a validação desta requisição, pois existe(m) produto(s) com o centro de custo autorizado."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                    TBAbrir.Close
                Else
                    'Verifica se a empresa exige centro de custo
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and CC_obrigatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select IDrequisicao from Requisicao_materiais_lista where IDrequisicao = " & .ListItems(InitFor) & " and ID_CC IS NULL", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            USMsgBox ("Não é permitido validar esta requisição, pois o(s) produto(s) precisa(m) ter centro de custo cadastrado."), vbExclamation, "CAPRIND v5.0"
                            TBAbrir.Close
                            TBFI.Close
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                        TBAbrir.Close
                    End If
                    TBFI.Close
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

Private Sub Lista_req_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_req.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Requisicao_materiais where id = " & Lista_req.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
End If
CodigoLista = Lista_req.SelectedItem.index

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
        If Lista_req.Visible = True Then Lista_req.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        If FunVerifProsseguir = False Then Exit Sub
        Lista.SetFocus
        ProcCarregaLista_Item
    Case 2:
        Cmb_empresa.Visible = False
        If FunVerifProsseguir = False Then Exit Sub
        ProcCarregalistaCentroCusto
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifProsseguir() As Boolean
On Error GoTo tratar_erro

FunVerifProsseguir = True
If Novo_requisicao_material = True Then
    USMsgBox ("Salve a requisição antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    FunVerifProsseguir = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcCarregalistaCentroCusto()
On Error GoTo tratar_erro

If CC_RM = True Then TextoFiltro = " and RML.ID_CC = " & frmRelatorios_Custos_Prev_Real.ID_CC Else TextoFiltro = ""

valor = 0
Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "SELECT SUM(EM.VlrTotal) as valortotal, US.Codigo, US.Setor from ((Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RM.ID = RML.IDRequisicao) INNER JOIN Usuarios_Setor US ON RML.ID_CC = US.ID) LEFT JOIN Estoque_movimentacao EM ON EM.ID_prod_RM = RML.IDLista and EM.Documento = '" & txtrequisicao & "' where RM.ID = " & txtId & TextoFiltro & " Group By US.Codigo, US.Setor", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista(2).Min = 0
    PBLista(2).Max = TBLISTA.RecordCount
    PBLista(2).Value = 1
    Contador = 0
    With Lista2.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!CODIGO & " - " & TBLISTA!Setor
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!ValorTotal), "0,00", Format(TBLISTA!ValorTotal, "###,##0.00"))
            valor = valor + IIf(IsNull(TBLISTA!ValorTotal), 0, TBLISTA!ValorTotal)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista(2).Value = Contador
        Loop
        .Add , ""
        .Add , , "VALOR TOTAL:"
        .Item(.Count).SubItems(1) = Format(valor, "###,##0.00")
    End With
End If
TBLISTA.Close



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
Exit Sub
End Sub

Private Sub Txt_ordem_Change()
On Error GoTo tratar_erro

If Txt_ordem <> "" Then
    VerifNumero = tTxt_ordem
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ordem = ""
        Txt_ordem.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

txtdescricao.Text = ""
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

Lista_req.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_Localizar_Requisicao = "" Then Exit Sub
Set TBLISTA_Requisicao = CreateObject("adodb.recordset")
'Debug.print StrSql_Localizar_Requisicao
TBLISTA_Requisicao.Open StrSql_Localizar_Requisicao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Requisicao.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista_req.ListItems.Clear
TBLISTA_Requisicao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Requisicao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Requisicao.PageSize
ContadorReg = 1

PBLista(0).Min = 0
PBLista(0).Max = FunVerifMaxPBListaPaginacao(TBLISTA_Requisicao.RecordCount - IIf(Pagina > 1, (TBLISTA_Requisicao.PageSize * (Pagina - 1)), 0), TBLISTA_Requisicao.PageSize)
PBLista(0).Value = 1
Contador = 0
Do While TBLISTA_Requisicao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_req.ListItems
        .Add , , TBLISTA_Requisicao!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Requisicao!Empresa), "", TBLISTA_Requisicao!Empresa)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Requisicao!requisicao), "", TBLISTA_Requisicao!requisicao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Requisicao!Data), "", Format(TBLISTA_Requisicao!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Requisicao!Responsavel), "", TBLISTA_Requisicao!Responsavel)
        IDRequisicao = TBLISTA_Requisicao!ID
         ProcVerificaStatusRM
'        If IsNull(TBLISTA_Requisicao!status) = False Then
            If status = "ABERTA" Then .Item(.Count).SubItems(5) = "Aberta"
            If status = "RETIRADA" Then .Item(.Count).SubItems(5) = "Retirada"
            If status = "PARCIAL" Then .Item(.Count).SubItems(5) = "Retirada parcial"
        'End If
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Requisicao!DtValidacao) = False, "Sim", "Não")
    End With
    TBLISTA_Requisicao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista(0).Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Requisicao.RecordCount
If TBLISTA_Requisicao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Requisicao.PageCount
ElseIf TBLISTA_Requisicao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Requisicao.PageCount & " de: " & TBLISTA_Requisicao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Requisicao.AbsolutePage - 1 & " de: " & TBLISTA_Requisicao.PageCount
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

Private Sub txtQR_Change()
On Error GoTo tratar_erro

If txtQR.Text <> "" Then
    VerifNumero = txtQR.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQR.Text = ""
        txtQR.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQR_LostFocus()
On Error GoTo tratar_erro

txtQR = Format(txtQR, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtrequisicao_Change()
On Error GoTo tratar_erro

If Novo_requisicao_material = True Then
VerifNReq:
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select ID from Requisicao_materiais where requisicao = '" & txtrequisicao & "' and ID <> " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        Numero = ReturnNumbersOnly(Left(txtrequisicao, Len(txtrequisicao) - 3)) + 1
        Ano = Right(Year(Date), 2)
        NumeroReq = FunTamanhoTextoZeroEsq(Numero, 5) & "/" & Ano
        txtrequisicao = "RM-" & NumeroReq
        GoTo VerifNReq
    End If
    TBTempo.Close
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
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcCopiar
    Case 9: ProcValidarRegistros Lista_req, "Estoque/Requisição de materiais"
    Case 10: ProcAtualizar
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
    Case 1: ProcNovoProduto
    Case 2: ProcSalvarProduto
    Case 3: ProcExcluirProduto
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: procCentroDeCusto
    Case 8: procAutorizar
    'Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCentroDeCusto()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista.ListItems.Count = 0 Then Exit Sub

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de alterar o centro de custo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmRequisicao_materiais_CentroDeCusto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCarrega_CentroDeCusto()
On Error GoTo tratar_erro

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select A.* from Acessos A INNER JOIN Usuarios U ON A.IDUsuario = U.IDUsuario where U.Usuario = '" & txtResponsavel & "' and A.Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    ProcCarregaComboSetor Cmb_centro, "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and setor is not null and DtBloq IS NULL and (Consolidacao = 'False' or Consolidacao is null)", txtdesenho, False, True, False, "", True, False
Else
    ProcCarregaComboSetor Cmb_centro, "US.DtBloq Is Null and US.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), txtdesenho, False, True, False, txtResponsavel, True, False
End If
TBAcessos.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCarregaDados_Itens()
On Error GoTo tratar_erro

txtData_item.Text = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel_item = TBAbrir!Responsavel
txtdesenho.Text = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
If IsNull(TBAbrir!status) = False Then
    txtstatus_item.Text = Lista.SelectedItem.SubItems(8)
    If txtstatus_item = "Requisitado" Then Framelista.Enabled = True Else Framelista.Enabled = False
Else
    Framelista.Enabled = True
End If
txtdescricao.Text = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then cmbun = TBAbrir!Un
If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
Txt_ordem = IIf(IsNull(TBAbrir!Ordem), "", TBAbrir!Ordem)
txtQE = Format(FunVerificaQtdeEstoque(txtdesenho, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
txtQR.Text = IIf(IsNull(TBAbrir!Quant), "", Format(TBAbrir!Quant, "###,##0.0000"))

ProcVerifValorTotalSaida TBAbrir!IDlista
txtvlrUnit = Format(valor / TBAbrir!Quant, "###,##0.0000000000")
txtvlrTotal = Format(valor, "###,##0.00")

procCarrega_CentroDeCusto
If IsNull(TBAbrir!ID_CC) = False And TBAbrir!ID_CC <> "" Then
    NomeCampo = "o centro de custo"
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Codigo, Setor, DtBloq, ID from Usuarios_setor where ID = " & TBAbrir!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
            If IsNull(TBFI!DtBloq) = False Then
                Cmb_centro.AddItem TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
                Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
            End If
            Cmb_centro = TBFI!CODIGO & " - " & TBFI!Setor
        Else
            If IsNull(TBFI!DtBloq) = False Then
                Cmb_centro.AddItem IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
                Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
            End If
            Cmb_centro = TBFI!Setor
        End If
    End If
    TBFI.Close
End If

1:
    txtData_Autorizacao = IIf(IsNull(TBAbrir!Data_autorizacao), "", Format(TBAbrir!Data_autorizacao, "dd/mm/yy"))
    txtAutorizado = IIf(IsNull(TBAbrir!Autorizado), "", TBAbrir!Autorizado)
    txtObs_item.Text = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste produto."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procAutorizar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista.ListItems.Count = 0 Then Exit Sub

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de autorizar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmRequisicao_materiais_Autorizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
