VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcompras_reqcot 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Compras - Cotação"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   100
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
      ItemData        =   "frmcompras_reqcot.frx":0000
      Left            =   260
      List            =   "frmcompras_reqcot.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      ToolTipText     =   "Empresa."
      Top             =   1695
      Visible         =   0   'False
      Width           =   4500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Carteira de cotação"
      TabPicture(0)   =   "frmcompras_reqcot.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Dados da cotação"
      TabPicture(1)   =   "frmcompras_reqcot.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar3"
      Tab(1).Control(1)=   "lista_cot"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "txtidcot"
      Tab(1).Control(4)=   "Frame11"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Lista de fornecedores"
      TabPicture(2)   =   "frmcompras_reqcot.frx":003C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "USToolBar4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Lista_itens"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lista_forn"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtIDListaForn"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Lista de produtos/serviços"
      TabPicture(3)   =   "frmcompras_reqcot.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab3"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   170
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
            Index           =   2
            Left            =   9540
            TabIndex        =   43
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
            Index           =   2
            Left            =   3780
            TabIndex        =   42
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Index           =   2
            Left            =   11760
            TabIndex        =   47
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmcompras_reqcot.frx":0074
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
            Index           =   2
            Left            =   11220
            TabIndex        =   46
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmcompras_reqcot.frx":3818
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
            Index           =   2
            Left            =   10110
            TabIndex        =   44
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
            Index           =   2
            Left            =   10680
            TabIndex        =   45
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmcompras_reqcot.frx":7321
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
            Index           =   2
            Left            =   12300
            TabIndex        =   48
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmcompras_reqcot.frx":B410
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4410
            TabIndex        =   181
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   173
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   13050
            TabIndex        =   172
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3090
            TabIndex        =   171
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox txtidcot 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Left            =   -72300
         MaxLength       =   50
         MouseIcon       =   "frmcompras_reqcot.frx":EC9C
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Text            =   "0"
         Top             =   4290
         Visible         =   0   'False
         Width           =   1035
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
         Height          =   1755
         Left            =   -74925
         TabIndex        =   110
         Top             =   1305
         Width           =   15225
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
            Left            =   6795
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1425
         End
         Begin VB.TextBox txtrequisitante 
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
            Left            =   8235
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2880
         End
         Begin VB.TextBox cmbsetor 
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
            Left            =   11130
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Setor."
            Top             =   390
            Width           =   2595
         End
         Begin VB.TextBox txtidcotacao 
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
            Left            =   4680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Nº da cotação."
            Top             =   390
            Width           =   1305
         End
         Begin VB.TextBox txtdataemissao 
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
            Left            =   6000
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   780
         End
         Begin VB.TextBox txtobservacao 
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
            Height          =   615
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            ToolTipText     =   "Observações."
            Top             =   990
            Width           =   14865
         End
         Begin MSMask.MaskEdBox Txt_data_validade 
            Height          =   315
            Left            =   13740
            TabIndex        =   39
            ToolTipText     =   "Data de validade."
            Top             =   390
            Width           =   990
            _ExtentX        =   1746
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. validade"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   13800
            TabIndex        =   187
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   12232
            TabIndex        =   186
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   6218
            TabIndex        =   185
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
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
            Index           =   0
            Left            =   7230
            TabIndex        =   115
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   9218
            TabIndex        =   114
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº da cotação"
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
            Index           =   35
            Left            =   4740
            TabIndex        =   113
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   7140
            TabIndex        =   112
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label6 
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
            Left            =   2130
            TabIndex        =   111
            Top             =   180
            Width           =   735
         End
         Begin VB.Image Img_calendario_validade 
            Height          =   360
            Left            =   14730
            Picture         =   "frmcompras_reqcot.frx":EFA6
            Stretch         =   -1  'True
            ToolTipText     =   "Abrir calendário."
            Top             =   360
            Width           =   330
         End
      End
      Begin VB.TextBox txtIDListaForn 
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
         Left            =   4410
         MaxLength       =   50
         MouseIcon       =   "frmcompras_reqcot.frx":F429
         MousePointer    =   99  'Custom
         TabIndex        =   108
         Text            =   "0"
         ToolTipText     =   "ID."
         Top             =   3840
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   75
         TabIndex        =   101
         Top             =   1305
         Width           =   15225
         Begin VB.TextBox txtidforn 
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
            MaxLength       =   50
            MouseIcon       =   "frmcompras_reqcot.frx":F733
            MousePointer    =   99  'Custom
            TabIndex        =   49
            ToolTipText     =   "Código do fornecedor."
            Top             =   390
            Width           =   735
         End
         Begin VB.CommandButton cmdForn 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9180
            Picture         =   "frmcompras_reqcot.frx":FA3D
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Localizar fornecedores."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtcondpagtoforn 
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
            Left            =   4380
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Condição de pagamento."
            Top             =   990
            Width           =   8805
         End
         Begin VB.TextBox txtforn 
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
            Left            =   930
            Locked          =   -1  'True
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "Fornecedor."
            Top             =   390
            Width           =   8235
         End
         Begin VB.TextBox txtcontatoforn 
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
            Left            =   9570
            Locked          =   -1  'True
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Contato."
            Top             =   390
            Width           =   5115
         End
         Begin VB.TextBox txtfaxforn 
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
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   55
            ToolTipText     =   "Fax."
            Top             =   990
            Width           =   2085
         End
         Begin VB.TextBox txttelforn 
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
            MaxLength       =   30
            TabIndex        =   54
            ToolTipText     =   "Telefone."
            Top             =   990
            Width           =   2085
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Frete"
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   13590
            TabIndex        =   102
            Top             =   810
            Width           =   1455
            Begin VB.OptionButton Chkcifforn 
               BackColor       =   &H00E0E0E0&
               Caption         =   "CIF"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   180
               TabIndex        =   58
               Top             =   210
               Width           =   585
            End
            Begin VB.OptionButton Chkfobforn 
               BackColor       =   &H00E0E0E0&
               Caption         =   "FOB"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   750
               TabIndex        =   59
               Top             =   210
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdconforn 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   13185
            Picture         =   "frmcompras_reqcot.frx":FB3F
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Localizar condições de pagamento."
            Top             =   990
            Width           =   315
         End
         Begin VB.CommandButton cmdcontatosforn 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14715
            Picture         =   "frmcompras_reqcot.frx":FC41
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Localizar contatos."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condição de pagamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   7912
            TabIndex        =   107
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contato"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   11835
            TabIndex        =   106
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   3187
            TabIndex        =   105
            Top             =   780
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   907
            TabIndex        =   104
            Top             =   780
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   4635
            TabIndex        =   103
            Top             =   180
            Width           =   825
         End
      End
      Begin MSComctlLib.ListView lista_forn 
         Height          =   4080
         Left            =   75
         TabIndex        =   60
         Top             =   2775
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   7197
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
            Size            =   6.75
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
            Text            =   "Fornecedor"
            Object.Width           =   10998
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Contato"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Telefone"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Fax"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Cond. de pagamento"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "CIF"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "FOB"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "IDForn"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_itens 
         Height          =   2835
         Left            =   75
         TabIndex        =   61
         Top             =   6870
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   5001
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
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   24
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "IDLista"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº solicitação"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   3589
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1411
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
            Text            =   "Qtde. com."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Valor unit."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Desc. (%)"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Vlr. desc."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Valor unit. c/ desc."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "N"
            Text            =   "IPI (%)"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Object.Tag             =   "N"
            Text            =   "ICMS (%)"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Object.Tag             =   "N"
            Text            =   "Vlr. IPI"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Object.Tag             =   "N"
            Text            =   "Vlr. total"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   17
            Object.Tag             =   "D"
            Text            =   "Prazo entr."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Object.Tag             =   "T"
            Text            =   "Detalhe"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   22
            Object.Tag             =   "T"
            Text            =   "Aprov."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Object.Tag             =   "N"
            Text            =   "ID tabela forn"
            Object.Width           =   0
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   75
         TabIndex        =   109
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
         ButtonCaption7  =   "Adicionar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Adicionar o(s) produto(s)/serviço(s) ao(s) fornecedor(es) (F7)"
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
         ButtonLeft7     =   309
         ButtonTop7      =   2
         ButtonWidth7    =   52
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Aprovar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Aprovar (F8)"
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
         ButtonLeft8     =   363
         ButtonTop8      =   2
         ButtonWidth8    =   47
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Não aprovar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Não aprovar (F9)"
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
         ButtonLeft9     =   412
         ButtonTop9      =   2
         ButtonWidth9    =   68
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Cancelar"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Cancelar aprovação e não aprovação (F10)"
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
         ButtonLeft10    =   482
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
         ButtonLeft11    =   534
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
         ButtonLeft12    =   538
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
         ButtonLeft13    =   581
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
         ButtonLeft14    =   613
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   13950
            Top             =   180
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmcompras_reqcot.frx":FD43
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView lista_cot 
         Height          =   6015
         Left            =   -74925
         TabIndex        =   41
         Top             =   3075
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cotação"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Requisitante"
            Object.Width           =   6027
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Setor"
            Object.Width           =   6027
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "IDempresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "D"
            Text            =   "Dt. validade"
            Object.Width           =   2117
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   116
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
         ButtonCaption4  =   "Cancelar"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Cancelar (F4)"
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
         ButtonWidth4    =   57
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
         ButtonLeft5     =   183
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
         ButtonLeft6     =   245
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
         ButtonLeft7     =   302
         ButtonTop7      =   2
         ButtonWidth7    =   55
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Liberar/cancelar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Liberar/cancelar liberação (F8)"
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
         ButtonLeft8     =   359
         ButtonTop8      =   2
         ButtonWidth8    =   85
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Emitir PC"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Emitir/verificar pedido(s) de compra (F10)"
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
         ButtonLeft9     =   446
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
         ButtonLeft10    =   498
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
         ButtonLeft11    =   550
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
         ButtonLeft12    =   554
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
         ButtonLeft13    =   597
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
         ButtonLeft14    =   629
         ButtonTop14     =   2
         ButtonWidth14   =   24
         ButtonHeight14  =   24
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   14490
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmcompras_reqcot.frx":17F29
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   9750
         Left            =   -74925
         TabIndex        =   124
         Top             =   300
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   17198
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
         TabCaption(0)   =   "Dados dos produtos/serviços"
         TabPicture(0)   =   "frmcompras_reqcot.frx":2016C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "USToolBar5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Lista_itens1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtidlista"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtiditem"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame9"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Txt_ID_tabela_forn"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Txt_quantidade_PC"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Empenhos"
         TabPicture(1)   =   "frmcompras_reqcot.frx":20188
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista_empenhos"
         Tab(1).Control(1)=   "USToolBar6"
         Tab(1).Control(2)=   "Frame16"
         Tab(1).ControlCount=   3
         Begin VB.TextBox Txt_quantidade_PC 
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
            Left            =   6210
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   184
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade em peça."
            Top             =   6060
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   -74940
            TabIndex        =   155
            Top             =   8550
            Width           =   15075
            Begin VB.TextBox Txt_qtde_total_cotada 
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
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   96
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade total cotada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_emp 
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
               Left            =   11430
               Locked          =   -1  'True
               TabIndex        =   97
               TabStop         =   0   'False
               ToolTipText     =   "Quatidade total empenhada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_disp 
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
               Left            =   13320
               Locked          =   -1  'True
               TabIndex        =   98
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade disponível."
               Top             =   420
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "-                                       ="
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
               Index           =   26
               Left            =   11250
               TabIndex        =   157
               Top             =   480
               Width           =   1965
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. cotada           Qtde. empenhada          Qtde. disponível"
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
               Index           =   25
               Left            =   9870
               TabIndex        =   156
               Top             =   210
               Width           =   4875
            End
         End
         Begin VB.TextBox Txt_ID_tabela_forn 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   5160
            MaxLength       =   50
            MouseIcon       =   "frmcompras_reqcot.frx":201A4
            MousePointer    =   99  'Custom
            TabIndex        =   154
            Text            =   "0"
            Top             =   6060
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   60
            TabIndex        =   152
            Top             =   1305
            Width           =   15105
            Begin VB.TextBox txtidforn1 
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
               MaxLength       =   50
               MouseIcon       =   "frmcompras_reqcot.frx":204AE
               MousePointer    =   99  'Custom
               TabIndex        =   62
               TabStop         =   0   'False
               ToolTipText     =   "Código do fornecedor."
               Top             =   390
               Width           =   735
            End
            Begin VB.TextBox txtforn1 
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
               Left            =   930
               Locked          =   -1  'True
               MaxLength       =   255
               ScrollBars      =   2  'Vertical
               TabIndex        =   63
               TabStop         =   0   'False
               ToolTipText     =   "Fornecedor."
               Top             =   390
               Width           =   13995
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fornecedor"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   12
               Left            =   7575
               TabIndex        =   153
               Top             =   180
               Width           =   825
            End
         End
         Begin VB.TextBox txtiditem 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   4110
            MaxLength       =   50
            MouseIcon       =   "frmcompras_reqcot.frx":207B8
            MousePointer    =   99  'Custom
            TabIndex        =   151
            Text            =   "0"
            Top             =   6060
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtidlista 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Left            =   3060
            MaxLength       =   50
            MouseIcon       =   "frmcompras_reqcot.frx":20AC2
            MousePointer    =   99  'Custom
            TabIndex        =   150
            Text            =   "0"
            Top             =   6060
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Frame Frame4 
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
            Height          =   2985
            Left            =   60
            TabIndex        =   126
            Top             =   2175
            Width           =   15105
            Begin VB.ComboBox cmbReferencia 
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
               ItemData        =   "frmcompras_reqcot.frx":20DCC
               Left            =   2370
               List            =   "frmcompras_reqcot.frx":20DCE
               Sorted          =   -1  'True
               TabIndex        =   66
               ToolTipText     =   "Código de referencia."
               Top             =   390
               Width           =   2325
            End
            Begin VB.TextBox Txt_vlr_unit_ultima_compra 
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
               Left            =   13350
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   75
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário da última compra."
               Top             =   1020
               Width           =   1575
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1950
               Picture         =   "frmcompras_reqcot.frx":20DD0
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Visualizar arquivo."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox Txt_quantidade_est 
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
               Left            =   13845
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   71
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade da unidade de estoque."
               Top             =   390
               Width           =   1080
            End
            Begin VB.TextBox Txt_obs_fornecedor 
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
               Height          =   615
               Left            =   10380
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               ToolTipText     =   "Observações do fornecedor."
               Top             =   1610
               Width           =   3645
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
               ForeColor       =   &H00000000&
               Height          =   330
               ItemData        =   "frmcompras_reqcot.frx":21392
               Left            =   11895
               List            =   "frmcompras_reqcot.frx":21394
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   69
               TabStop         =   0   'False
               ToolTipText     =   "Unidade comercial."
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtdesconto 
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
               Left            =   3630
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   85
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   2550
               Width           =   1140
            End
            Begin VB.TextBox txtvalorunitariodesc 
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
               Left            =   6436
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   88
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   2550
               Width           =   1775
            End
            Begin VB.TextBox txtvalordesconto 
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
               Left            =   4780
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   87
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   2550
               Width           =   1635
            End
            Begin VB.TextBox txttotalforn 
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
               Left            =   13260
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   93
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   2550
               Width           =   1665
            End
            Begin VB.TextBox txtipi 
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
               Left            =   8214
               MaxLength       =   50
               TabIndex        =   89
               ToolTipText     =   "Valor de % IPI."
               Top             =   2550
               Width           =   705
            End
            Begin VB.TextBox txtIcms 
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
               Left            =   10720
               MaxLength       =   50
               TabIndex        =   91
               ToolTipText     =   "Valor de % ICMS."
               Top             =   2550
               Width           =   735
            End
            Begin VB.TextBox TxtvlrIpi 
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
               Left            =   8942
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   90
               TabStop         =   0   'False
               ToolTipText     =   "Valor de IPI."
               Top             =   2550
               Width           =   1775
            End
            Begin VB.TextBox TxtVlrIcms 
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
               Left            =   11478
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   92
               TabStop         =   0   'False
               ToolTipText     =   "Valor de ICMS."
               Top             =   2550
               Width           =   1775
            End
            Begin VB.TextBox txtDescricao_comercial 
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
               Height          =   615
               Left            =   180
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   76
               ToolTipText     =   "Descrição comercial."
               Top             =   1610
               Width           =   6525
            End
            Begin VB.CheckBox Chk_desc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desc. (%)"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3683
               TabIndex        =   84
               Top             =   2340
               Width           =   1035
            End
            Begin VB.CheckBox Chk_valor_desc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Vlr. desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5080
               TabIndex        =   86
               Top             =   2340
               Width           =   1035
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Aprovado"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   705
               Left            =   14010
               TabIndex        =   127
               Top             =   1500
               Width           =   915
               Begin VB.OptionButton Optnao 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Não"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   180
                  TabIndex        =   80
                  Top             =   450
                  Width           =   585
               End
               Begin VB.OptionButton Optsim 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Sim"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   180
                  TabIndex        =   79
                  Top             =   240
                  Width           =   555
               End
            End
            Begin VB.CommandButton cmdcalc_peso 
               BackColor       =   &H00C0C0C0&
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
               Left            =   3210
               Picture         =   "frmcompras_reqcot.frx":21396
               Style           =   1  'Graphical
               TabIndex        =   83
               ToolTipText     =   "Abrir calculadora para cálculo de peso."
               Top             =   2550
               Width           =   315
            End
            Begin VB.TextBox txtvalorunitforn 
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
               Left            =   1620
               MaxLength       =   50
               TabIndex        =   82
               ToolTipText     =   "Valor unitário."
               Top             =   2550
               Width           =   1575
            End
            Begin VB.TextBox Txt_fornecedor_aprovado 
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
               Left            =   7710
               Locked          =   -1  'True
               MaxLength       =   255
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               TabStop         =   0   'False
               ToolTipText     =   "Fornecedor aprovado."
               Top             =   1020
               Width           =   5625
            End
            Begin VB.TextBox Txt_quantidade 
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
               Left            =   12750
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   70
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade comercial."
               Top             =   390
               Width           =   1080
            End
            Begin VB.ComboBox Cmb_un 
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
               ItemData        =   "frmcompras_reqcot.frx":215FF
               Left            =   11070
               List            =   "frmcompras_reqcot.frx":21601
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   68
               TabStop         =   0   'False
               ToolTipText     =   "Unidade de estoque."
               Top             =   390
               Width           =   825
            End
            Begin VB.ComboBox Cmb_familia 
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
               Left            =   2520
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   73
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   1020
               Width           =   5190
            End
            Begin VB.TextBox txtStatusProd 
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
               TabIndex        =   72
               TabStop         =   0   'False
               ToolTipText     =   "Status."
               Top             =   1020
               Width           =   2325
            End
            Begin VB.TextBox txtDescricao_item 
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
               Left            =   4710
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   67
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   390
               Width           =   6345
            End
            Begin VB.TextBox txtDesenho 
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
               TabIndex        =   64
               TabStop         =   0   'False
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   1755
            End
            Begin VB.TextBox txtOBS 
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
               Height          =   615
               Left            =   6720
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   77
               ToolTipText     =   "Observações."
               Top             =   1610
               Width           =   3645
            End
            Begin MSMask.MaskEdBox txtPrazoentregaforn 
               Height          =   315
               Left            =   180
               TabIndex        =   81
               ToolTipText     =   "Prazo de entrega."
               Top             =   2550
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   0
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Código referência"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   17
               Left            =   2895
               TabIndex        =   149
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. última compra"
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
               Index           =   18
               Left            =   13365
               TabIndex        =   148
               Top             =   810
               Width           =   1545
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. un. est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   17
               Left            =   13868
               TabIndex        =   147
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observações fornecedor"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   16
               Left            =   11310
               TabIndex        =   146
               Top             =   1410
               Width           =   1785
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Un. com."
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   12000
               TabIndex        =   145
               Top             =   180
               Width           =   645
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição comercial"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   15
               Left            =   2745
               TabIndex        =   144
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Image imgCalendario 
               Height          =   360
               Left            =   1230
               Picture         =   "frmcompras_reqcot.frx":21603
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   2520
               Width           =   330
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   11975
               TabIndex        =   143
               Top             =   2340
               Width           =   780
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor IPI"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9522
               TabIndex        =   142
               Top             =   2340
               Width           =   615
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "% ICMS"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   10795
               TabIndex        =   141
               Top             =   2340
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "% IPI"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   8356
               TabIndex        =   140
               Top             =   2340
               Width           =   420
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor unitário"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   11
               Left            =   1935
               TabIndex        =   139
               Top             =   2340
               Width           =   945
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor total"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   13725
               TabIndex        =   138
               Top             =   2340
               Width           =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo entrega"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   9
               Left            =   202
               TabIndex        =   137
               Top             =   2340
               Width           =   1020
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vlr. unit. c/ desc."
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6708
               TabIndex        =   136
               Top             =   2340
               Width           =   1230
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fornecedor aprovado"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   21
               Left            =   9742
               TabIndex        =   135
               Top             =   810
               Width           =   1560
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. com."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   14
               Left            =   12885
               TabIndex        =   134
               Top             =   180
               Width           =   810
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Un. est."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   13
               Left            =   11190
               TabIndex        =   133
               Top             =   180
               Width           =   585
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Família"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   20
               Left            =   4860
               TabIndex        =   132
               Top             =   810
               Width           =   480
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
               Left            =   1065
               TabIndex        =   131
               Top             =   810
               Width           =   555
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   7537
               TabIndex        =   130
               Top             =   180
               Width           =   690
            End
            Begin VB.Label Label2 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   547
               TabIndex        =   129
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observações"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   8070
               TabIndex        =   128
               Top             =   1410
               Width           =   945
            End
         End
         Begin MSComctlLib.ListView Lista_itens1 
            Height          =   4530
            Left            =   60
            TabIndex        =   94
            Top             =   5175
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   7990
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   24
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "IDLista"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Nº solicitação"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   3589
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Un. est."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Un. com."
               Object.Width           =   1411
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
               Text            =   "Qtde. com."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Valor unit."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Desc. (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Vlr. desc."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "Vlr. unit. c/ desc."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   13
               Object.Tag             =   "N"
               Text            =   "IPI (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   14
               Object.Tag             =   "N"
               Text            =   "ICMS (%)"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   15
               Object.Tag             =   "N"
               Text            =   "Vlr. IPI"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   16
               Object.Tag             =   "N"
               Text            =   "Vlr. total"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   17
               Object.Tag             =   "D"
               Text            =   "Prazo entr."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Object.Tag             =   "T"
               Text            =   "Detalhe"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   19
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   20
               Object.Tag             =   "N"
               Text            =   "Ordem"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   21
               Object.Tag             =   "N"
               Text            =   "OS"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   22
               Object.Tag             =   "T"
               Text            =   "Aprov."
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   23
               Object.Tag             =   "N"
               Text            =   "ID tabela forn"
               Object.Width           =   0
            EndProperty
         End
         Begin DrawSuite2022.USToolBar USToolBar6 
            Height          =   975
            Left            =   -74940
            TabIndex        =   158
            Top             =   330
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   1720
            ButtonCount     =   7
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
            ButtonLeft2     =   40
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
            ButtonLeft3     =   87
            ButtonTop3      =   2
            ButtonWidth3    =   60
            ButtonHeight3   =   21
            ButtonUseMaskColor3=   0   'False
            ButtonEnabled4  =   0   'False
            ButtonIconSize4 =   32
            ButtonAlignment4=   2
            ButtonType4     =   1
            ButtonStyle4    =   -1
            BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonState4    =   -1
            ButtonLeft4     =   149
            ButtonTop4      =   4
            ButtonWidth4    =   2
            ButtonHeight4   =   54
            ButtonCaption5  =   "Ajuda"
            ButtonEnabled5  =   0   'False
            ButtonIconSize5 =   32
            ButtonToolTipText5=   "Ajuda (F1)"
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
            ButtonLeft5     =   153
            ButtonTop5      =   2
            ButtonWidth5    =   41
            ButtonHeight5   =   21
            ButtonUseMaskColor5=   0   'False
            ButtonCaption6  =   "Sair"
            ButtonEnabled6  =   0   'False
            ButtonIconSize6 =   32
            ButtonToolTipText6=   "Sair (Esc)"
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
            ButtonLeft6     =   196
            ButtonTop6      =   2
            ButtonWidth6    =   30
            ButtonHeight6   =   21
            ButtonUseMaskColor6=   0   'False
            ButtonEnabled7  =   0   'False
            ButtonIconSize7 =   32
            ButtonKey7      =   "7"
            ButtonAlignment7=   2
            BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonState7    =   5
            ButtonLeft7     =   228
            ButtonTop7      =   2
            ButtonWidth7    =   24
            ButtonHeight7   =   24
            ButtonUseMaskColor7=   0   'False
            Begin DrawSuite2022.USImageList USImageList6 
               Left            =   14310
               Top             =   240
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmcompras_reqcot.frx":21A86
               Count           =   1
            End
         End
         Begin MSComctlLib.ListView Lista_empenhos 
            Height          =   7215
            Left            =   -74940
            TabIndex        =   95
            Top             =   1320
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   12726
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   17
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Cód. cart."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "N"
               Text            =   "Ped. int./SPR"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   2297
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Cod. ref."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   2914
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Qtde. rec."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Object.Tag             =   "T"
               Text            =   "Ped. cliente"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Object.Tag             =   "T"
               Text            =   "N. item"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   1764
            EndProperty
         End
         Begin DrawSuite2022.USToolBar USToolBar5 
            Height          =   975
            Left            =   60
            TabIndex        =   159
            Top             =   330
            Width           =   15105
            _ExtentX        =   26644
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
            ButtonToolTipText4=   "Fornecedor anterior."
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
            ButtonToolTipText5=   "Próximo fornecedor."
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
            ButtonCaption6  =   "Excluir todos"
            ButtonEnabled6  =   0   'False
            ButtonIconSize6 =   32
            ButtonToolTipText6=   "Excluir produto/serviço de todos os fornecedores (F7)"
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
            ButtonLeft6     =   271
            ButtonTop6      =   2
            ButtonWidth6    =   69
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
            ButtonLeft7     =   342
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
            ButtonLeft8     =   346
            ButtonTop8      =   2
            ButtonWidth8    =   41
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonLeft9     =   389
            ButtonTop9      =   2
            ButtonWidth9    =   30
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
            ButtonLeft10    =   421
            ButtonTop10     =   2
            ButtonWidth10   =   24
            ButtonHeight10  =   24
            ButtonUseMaskColor10=   0   'False
            Begin DrawSuite2022.USImageList USImageList5 
               Left            =   14310
               Top             =   240
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmcompras_reqcot.frx":25349
               Count           =   1
            End
         End
      End
      Begin VB.Frame Frame3 
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
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74860
         TabIndex        =   160
         Top             =   1620
         Width           =   5805
         Begin VB.ComboBox Cmb_empresa_carteira 
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
            ItemData        =   "frmcompras_reqcot.frx":2A914
            Left            =   210
            List            =   "frmcompras_reqcot.frx":2A91B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   390
            Width           =   5415
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
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
            Left            =   2535
            TabIndex        =   161
            Top             =   180
            Width           =   765
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74880
         TabIndex        =   178
         Top             =   630
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   1720
         ButtonCount     =   6
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
         ButtonCaption2  =   "Gerar cot."
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Gerar cotação (F3)"
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
         ButtonWidth2    =   64
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonAlignment3=   2
         ButtonType3     =   1
         ButtonStyle3    =   -1
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   -1
         ButtonLeft3     =   112
         ButtonTop3      =   4
         ButtonWidth3    =   2
         ButtonHeight3   =   54
         ButtonCaption4  =   "Ajuda"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Ajuda (F1)"
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
         ButtonLeft4     =   116
         ButtonTop4      =   2
         ButtonWidth4    =   41
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Sair"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Sair (Esc)"
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
         ButtonLeft5     =   159
         ButtonTop5      =   2
         ButtonWidth5    =   30
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonKey6      =   "6"
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState6    =   5
         ButtonLeft6     =   191
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   5520
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmcompras_reqcot.frx":2A92C
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   9750
         Left            =   -74925
         TabIndex        =   118
         Top             =   300
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   17198
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
         TabCaption(0)   =   "Necessidade"
         TabPicture(0)   =   "frmcompras_reqcot.frx":2D479
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ListaNecessidade"
         Tab(0).Control(1)=   "Frame1(23)"
         Tab(0).Control(2)=   "Frame15"
         Tab(0).Control(3)=   "Frame10"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Solicitação"
         TabPicture(1)   =   "frmcompras_reqcot.frx":2D495
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Lista_solicitados"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame7"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame1(2)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Index           =   2
            Left            =   5880
            TabIndex        =   174
            Top             =   1320
            Width           =   9285
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   555
               Index           =   3
               Left            =   2640
               TabIndex        =   175
               Top             =   165
               Width           =   2355
               Begin VB.OptionButton optIgual_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Igual"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Left            =   1290
                  TabIndex        =   182
                  Top             =   330
                  Width           =   645
               End
               Begin VB.OptionButton Optmeio_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Meio frase"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Left            =   180
                  TabIndex        =   31
                  Top             =   330
                  Width           =   975
               End
               Begin VB.OptionButton Optinicio_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Início frase"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Left            =   180
                  TabIndex        =   30
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1065
               End
               Begin VB.OptionButton OptFim_sol 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Fim frase"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Left            =   1290
                  TabIndex        =   32
                  Top             =   120
                  Width           =   915
               End
            End
            Begin VB.ComboBox cmbfiltrarpor_sol 
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
               ItemData        =   "frmcompras_reqcot.frx":2D4B1
               Left            =   180
               List            =   "frmcompras_reqcot.frx":2D4CD
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   18
               ToolTipText     =   "Opções para filtro."
               Top             =   390
               Width           =   2445
            End
            Begin VB.TextBox txtTexto_sol 
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
               TabIndex        =   19
               ToolTipText     =   "Texto para pesquisa."
               Top             =   390
               Width           =   4095
            End
            Begin MSComCtl2.DTPicker Txtprazo_sol 
               Height          =   315
               Left            =   5010
               TabIndex        =   21
               ToolTipText     =   "Texto para pesquisa."
               Top             =   390
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
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
               Format          =   198574081
               CurrentDate     =   39057
            End
            Begin VB.ComboBox cmbTexto_sol 
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
               ItemData        =   "frmcompras_reqcot.frx":2D53C
               Left            =   5010
               List            =   "frmcompras_reqcot.frx":2D53E
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   20
               ToolTipText     =   "Familia."
               Top             =   390
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texto para pesquisa"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   6315
               TabIndex        =   177
               Top             =   180
               Width           =   1485
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filtrar por"
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
               Left            =   1057
               TabIndex        =   176
               Top             =   180
               Width           =   840
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   75
            TabIndex        =   166
            Top             =   8790
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
               Index           =   1
               Left            =   9540
               TabIndex        =   24
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
               Index           =   1
               Left            =   3780
               TabIndex        =   23
               Text            =   "30"
               ToolTipText     =   "Número de registros por página."
               Top             =   180
               Width           =   555
            End
            Begin DrawSuite2022.USButton cmdPagProx 
               Height          =   315
               Index           =   1
               Left            =   11760
               TabIndex        =   28
               ToolTipText     =   "Próxima página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":2D540
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
               Index           =   1
               Left            =   11220
               TabIndex        =   27
               ToolTipText     =   "Página anterior."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":30CE4
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
               Index           =   1
               Left            =   10110
               TabIndex        =   25
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
               Index           =   1
               Left            =   10680
               TabIndex        =   26
               ToolTipText     =   "Primeira página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":347ED
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
               Index           =   1
               Left            =   12300
               TabIndex        =   29
               ToolTipText     =   "Última página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":388DC
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4410
               TabIndex        =   179
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label lblRegistros 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de registros: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   169
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label lblPaginas 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Página: 0 de: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   13050
               TabIndex        =   168
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Carregar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3090
               TabIndex        =   167
               Top             =   240
               Width           =   645
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   -74925
            TabIndex        =   162
            Top             =   8790
            Width           =   15105
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
               Index           =   0
               Left            =   3780
               TabIndex        =   8
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
               Index           =   0
               Left            =   9540
               TabIndex        =   9
               ToolTipText     =   "Número da página."
               Top             =   180
               Width           =   555
            End
            Begin DrawSuite2022.USButton cmdPagProx 
               Height          =   315
               Index           =   0
               Left            =   11760
               TabIndex        =   13
               ToolTipText     =   "Próxima página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":3C168
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
               Index           =   0
               Left            =   11220
               TabIndex        =   12
               ToolTipText     =   "Página anterior."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":3F90C
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
               Index           =   0
               Left            =   10110
               TabIndex        =   10
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
               Index           =   0
               Left            =   10680
               TabIndex        =   11
               ToolTipText     =   "Primeira página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":43415
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
               Index           =   0
               Left            =   12300
               TabIndex        =   14
               ToolTipText     =   "Última página."
               Top             =   180
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   556
               DibPicture      =   "frmcompras_reqcot.frx":47504
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
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "registros por página"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4410
               TabIndex        =   180
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Carregar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3090
               TabIndex        =   165
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lblPaginas 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Página: 0 de: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   13050
               TabIndex        =   164
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lblRegistros 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de registros: 0"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   163
               Top             =   240
               Width           =   1275
            End
         End
         Begin VB.Frame Frame15 
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
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   -69120
            TabIndex        =   125
            Top             =   1320
            Width           =   9285
            Begin VB.OptionButton Opt_PCP 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Por PCP"
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
               Left            =   3630
               TabIndex        =   1
               Top             =   360
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton Opt_vendas 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Por vendas"
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
               Left            =   4740
               TabIndex        =   2
               Top             =   360
               Width           =   1245
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Index           =   23
            Left            =   -74925
            TabIndex        =   119
            Top             =   2190
            Width           =   15105
            Begin VB.ComboBox Cmb_filtrar 
               BackColor       =   &H00FFFFFF&
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
               Height          =   315
               ItemData        =   "frmcompras_reqcot.frx":4AD90
               Left            =   12630
               List            =   "frmcompras_reqcot.frx":4AD9A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               ToolTipText     =   "Tipo de necessidade."
               Top             =   390
               Width           =   2295
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   510
               Index           =   19
               Left            =   2640
               TabIndex        =   120
               Top             =   210
               Width           =   4695
               Begin VB.OptionButton optIgual_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Igual"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3870
                  TabIndex        =   183
                  Top             =   180
                  Width           =   705
               End
               Begin VB.OptionButton optFim_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Fim frase"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2720
                  TabIndex        =   17
                  Top             =   180
                  Width           =   1095
               End
               Begin VB.OptionButton optInicio_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Início frase"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   180
                  TabIndex        =   15
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   1245
               End
               Begin VB.OptionButton optMeio_necess 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Meio frase"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1480
                  TabIndex        =   16
                  Top             =   180
                  Width           =   1185
               End
            End
            Begin VB.ComboBox cmbfiltrarpor_necess 
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
               ItemData        =   "frmcompras_reqcot.frx":4ADC4
               Left            =   180
               List            =   "frmcompras_reqcot.frx":4ADE0
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   3
               ToolTipText     =   "Opções para filtro."
               Top             =   390
               Width           =   2415
            End
            Begin VB.TextBox txtTexto_necess 
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
               Left            =   7380
               TabIndex        =   4
               ToolTipText     =   "Texto para pesquisa."
               Top             =   390
               Width           =   5235
            End
            Begin VB.ComboBox cmbTexto_necess 
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
               ItemData        =   "frmcompras_reqcot.frx":4AE58
               Left            =   7380
               List            =   "frmcompras_reqcot.frx":4AE5A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   5
               ToolTipText     =   "Texto para pesquisa."
               Top             =   390
               Visible         =   0   'False
               Width           =   5235
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de necessidade"
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
               Left            =   12915
               TabIndex        =   123
               Top             =   180
               Width           =   1710
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Filtrar por"
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
               Index           =   24
               Left            =   967
               TabIndex        =   122
               Top             =   180
               Width           =   840
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texto para pesquisa"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   9255
               TabIndex        =   121
               Top             =   180
               Width           =   1485
            End
         End
         Begin MSComctlLib.ListView ListaNecessidade 
            Height          =   5700
            Left            =   -74925
            TabIndex        =   7
            Top             =   3060
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   10054
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
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   18106
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Un."
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Necessidade"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Necess. PÇ"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_solicitados 
            Height          =   6585
            Left            =   75
            TabIndex        =   22
            Top             =   2190
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
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Nº solicitação"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6730
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Un. est."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Un. com."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Quant. est."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Quant. com."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Detalhe"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Object.Tag             =   "D"
               Text            =   "Prazo entr."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "T"
               Text            =   "Obs."
               Object.Width           =   0
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmcompras_reqcot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fornecedor            As Variant 'OK
Dim Liberar                  As Boolean 'OK
Public Novo_Cotacao          As Boolean 'OK
Public Novo_Cotacao1         As Boolean 'OK
Public Novo_Cotacao2         As Boolean 'OK
Dim StrSql_Cotacao_Necessidade As String 'OK
Dim StrSql_Cotacao_Solicitacao As String 'OK
Public Sql_Cotacao_Localizar As String 'OK
Dim TBLISTA_Cotacao_Necessidade  As ADODB.Recordset 'OK
Dim TBLISTA_Cotacao_Solicitacao  As ADODB.Recordset 'OK
Dim TBLISTA_Compras_Cotacao  As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=iBxf8AMNGN8&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=7&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_desc_Click()
On Error GoTo tratar_erro

With txtDesconto
    If Chk_desc.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_valor_desc.Value = 0
        txtvalordesconto.Locked = True
        txtvalordesconto.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_valor_desc_Click()
On Error GoTo tratar_erro

With txtvalordesconto
    If Chk_valor_desc.Value = 1 Then
        .Locked = False
        .TabStop = True
        .SetFocus
        Chk_desc.Value = 0
        txtDesconto.Locked = True
        txtDesconto.TabStop = False
    Else
        .Locked = True
        .TabStop = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCamposCotacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCancelar_Aprovacao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtStatus = "APROVADA" Then
    USMsgBox ("Não é permitido cancelar o status do(s) produto(s)/serviço(s), pois a cotação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With Lista_itens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente cancelar o status desse(s) produto(s)/serviço(s) em todos os fornecedores?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "Update Cotacao_fornecedor Set naprovadoforn = 'False', aprovadoforn = 'False' where iditem = " & .ListItems.Item(InitFor).Text
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Cancelar status do produto/serviço"
            ID_documento = .ListItems.Item(InitFor).ListSubItems(1)
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) na lista antes de cancelar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcVerifPermissaoUsuario
    If Permitido = True Then
        Conexao.Execute "Update Compras_Cotacao Set Autorizado = NULL, Statuscotacao = 'COTANDO', dataliberada = NULL where id_cotacao = " & txtidcot
        txtStatus = "COTANDO"
    End If
    ProcCarregaListaItens
    ProcCarregaListaItens1
    ProcLimpaCamposItem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluirTodosForn()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If ProcVerifSatus("excluir este produto/serviço", True) = False Then Exit Sub

Permitido = False
With Lista_itens1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s) de todos os fornecedores?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select cotacao_fornecedor.id, cotacao_fornecedor.idforn, cotacao_fornecedor.iditem, Cotacao_item.ID, Cotacao_item.iditemlista FROM cotacao_fornecedor INNER JOIN Cotacao_item ON cotacao_fornecedor.iditem = Cotacao_item.ID where cotacao_fornecedor.idcot = " & txtidcot & " and Cotacao_item.ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    'Verifica se tem outro produto para este fornecedor
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select id from cotacao_fornecedor where idcot = " & txtidcot & " and idforn = " & TBFI!IDforn & " and IDitem <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        Conexao.Execute "DELETE from cotacao_fornecedor where idcot = " & txtidcot & " and idforn = " & TBFI!IDforn & " and IDitem = " & .ListItems(InitFor)
                    Else
                        TBFI!IDitem = 0
                        TBFI.Update
                    End If
                    Conexao.Execute "DELETE from CPLE from Compras_pedido_lista_empenhos CPLE INNER JOIN Compras_pedido_lista CPL ON CPL.IdLista = CPLE.IdLista where CPL.IdLista = " & TBFI!iditemlista & " and CPL.ID_Requisicao = 0"
                    Conexao.Execute "DELETE from Compras_pedido_lista where IdLista = " & TBFI!iditemlista & " and ID_Requisicao = 0"
                    Conexao.Execute "Update Compras_pedido_lista Set Status_item = 'REQUISIT.', ID_cotacao = 0 where IdLista = " & TBFI!iditemlista & " and ID_Requisicao <> 0"
                    TBFI.MoveNext
                Loop
                Conexao.Execute "DELETE from Cotacao_item where ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
            
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Excluir produto/serviço dos fornecedor"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("produto(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposItem
    ProcCarregaListaItens
    ProcCarregaListaItens1
    Frame4.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNao_aprovar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_itens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar para não aprovado este(s) produto(s)/serviço(s) em todos os fornecedores?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "Update Cotacao_fornecedor Set naprovadoforn = 'True', aprovadoforn = 'False' where iditem = " & .ListItems.Item(InitFor).Text
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Alterar produto/serviço para não aprovado"
            ID_documento = .ListItems.Item(InitFor).ListSubItems(1)
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) na lista antes de alterar para não aprovado."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaItens
    ProcCarregaListaItens1
    ProcLimpaCamposItem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtidcot = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_Cotacao order by Id_cotacao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Id_cotacao = " & txtidcot)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtidcot = TBLISTA!ID_cotacao
        Set TBCotacao = CreateObject("adodb.recordset")
        TBCotacao.Open "Select * from Compras_Cotacao where Id_cotacao = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCamposCotacao
        ProcLimpaCamposItem
        ProcLimpaCamposForn
        Lista_itens1.ListItems.Clear
        lista_forn.ListItems.Clear
        ProcCarregaDados
        ProcCarregaListaForn False
    Else
        USMsgBox ("Fim dos cadastros de cotação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Cotacao2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAnteriorForn()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Forn from Cotacao_fornecedor where idcot = " & txtidcot & " Group by forn", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    TBLISTA.Find ("Forn = '" & txtforn & "'")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Cotacao_fornecedor where idcot = " & txtidcot & " and Forn = '" & TBLISTA!forn & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcLimpaCamposForn
            ProcLimpaCamposItem
            ProcCarregaDadosForn
            ProcCarregaListaItens1
        End If
    Else
        USMsgBox ("Fim dos cadastros de fornecedores."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Cotacao2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAprovarForn()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_itens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente aprovar este(s) produto(s)/serviço(s) com o fornecedor " & txtforn & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from compras_pedido_lista where id_cotacao = " & txtidcot & " and IDlista = " & .ListItems.Item(InitFor).ListSubItems(1) & " and status_item <> 'Cancelado'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select * from cotacao_item where idcot = " & Cont & " and iditemlista = " & TBAbrir!IDlista & " order by id", Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        Set TBFornecedor = CreateObject("adodb.recordset")
                        TBFornecedor.Open "Select * from cotacao_fornecedor where idcot = " & Cont & " and iditem = " & TBItem!ID & " and Idforn = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFornecedor.EOF = False Then
                            Do While TBFornecedor.EOF = False
                                NomeCampo = ""
                                If TBFornecedor!contforn = "" Or IsNull(TBFornecedor!contforn) = True Then
                                    If NomeCampo <> "" Then NomeCampo = NomeCampo & ", contato" Else NomeCampo = "contato"
                                End If
                                If TBFornecedor!CondPagtoforn = "" Or IsNull(TBFornecedor!CondPagtoforn) = True Then
                                    If NomeCampo <> "" Then NomeCampo = NomeCampo & ", condição de pagamento" Else NomeCampo = "condição de pagamento"
                                End If
                                If TBFornecedor!CIFforn = False And TBFornecedor!FOBforn = False Then
                                    If NomeCampo <> "" Then NomeCampo = NomeCampo & ", frete" Else NomeCampo = "frete"
                                End If
                                If TBFornecedor!precounit = "" Or IsNull(TBFornecedor!precounit) = True Or TBFornecedor!precounit = "0" Then
                                    If NomeCampo <> "" Then NomeCampo = NomeCampo & ", valor unitário" Else NomeCampo = "valor unitário"
                                End If
                                If IsNull(TBFornecedor!prazoentregaforn) = True Then
                                    If NomeCampo <> "" Then NomeCampo = NomeCampo & ", prazo de entrega" Else NomeCampo = "prazo de entrega"
                                End If
                                If NomeCampo <> "" Then
                                    USMsgBox ("Não é permitido aprovar este(s) produto(s)/serviço(s), exite(m) algum(ns) dado(s) a ser(em) preenchido(s) para o produto/serviço: " & vbCrLf & TBAbrir!Desenho & " - " & TBAbrir!Descricao & vbCrLf & "Fornecedor: " & TBFornecedor!forn & vbCrLf & "Campos: " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
                                    TBAbrir.Close
                                    Exit Sub
                                End If
                                TBFornecedor.MoveNext
                            Loop
                        End If
                    End If
                    TBAbrir.MoveNext
                Loop
            End If
            
            Permitido = True
            Conexao.Execute "Update Cotacao_fornecedor Set naprovadoforn = 'False', aprovadoforn = 'True' where iditem = " & .ListItems.Item(InitFor) & " and idforn = " & lista_forn.SelectedItem.ListSubItems(8)
            Conexao.Execute "Update Cotacao_fornecedor Set naprovadoforn = 'True', aprovadoforn = 'False' where iditem = " & .ListItems.Item(InitFor) & " and idforn <> " & lista_forn.SelectedItem.ListSubItems(8)
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Aprovar produto/serviço"
            ID_documento = .ListItems.Item(InitFor).ListSubItems(1).Text
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) na lista antes de aprovar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("produto(s)/serviço(s) aprovado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    
    ProcVerifPermissaoUsuario
    If Permitido = True Then
        Conexao.Execute "Update Compras_Cotacao Set Autorizado = '" & pubUsuario & "', Statuscotacao = 'LIBERADA', dataliberada = '" & Format(Date, "Short Date") & "' where id_cotacao = " & txtidcot
        txtStatus = "LIBERADA"
    End If
    ProcCarregaListaItens
    ProcCarregaListaItens1
    ProcLimpaCamposItem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procSalvarForn()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtforn = "" Or txtIDforn = "" Or txtIDforn = "0" Then
    USMsgBox ("Informe o fornecedor antes de salvar."), vbExclamation, "CAPRIND v5.0"
    cmdForn_Click
    Exit Sub
End If
If Novo_Cotacao1 = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Cotacao_fornecedor where idcot = " & txtidcot & " and idforn =  " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Este fornecedor já está adicionado na lista."), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If
If ProcVerifSatus("salvar o fornecedor", True) = False Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from Cotacao_fornecedor where idcot = " & txtidcot & " and IDForn = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = True Then
    TBFornecedor.AddNew
    TBFornecedor!idcot = txtidcot
    TBFornecedor!IDforn = txtIDforn
    TBFornecedor!forn = txtforn
    TBFornecedor!contforn = txtcontatoforn
    TBFornecedor!telforn = txttelforn
    TBFornecedor!faxforn = txtfaxforn
    TBFornecedor!CondPagtoforn = txtcondpagtoforn
    If Chkcifforn.Value = True Then TBFornecedor!CIFforn = True Else TBFornecedor!CIFforn = False
    If Chkfobforn.Value = True Then TBFornecedor!FOBforn = True Else TBFornecedor!FOBforn = False
    TBFornecedor!aprovadoforn = False
    TBFornecedor!naprovadoforn = False
    TBFornecedor.Update
    txtIDListaForn = TBFornecedor!ID
Else
    If Chkcifforn.Value = True Then CIF1 = "CIFforn = 'True'" Else CIF1 = "CIFforn = 'False'"
    If Chkfobforn.Value = True Then FOB1 = "FOBforn = 'True'" Else FOB1 = "FOBforn = 'False'"
    Conexao.Execute "Update Cotacao_fornecedor Set contforn = '" & txtcontatoforn & "', telforn = '" & txttelforn & "', faxforn = '" & txtfaxforn & "', condpagtoforn = '" & txtcondpagtoforn & "', " & CIF1 & ", " & FOB1 & " where idcot = " & txtidcot & " and IDForn = " & txtIDforn
End If
TBFornecedor.Close

ProcCarregaListaForn False
If Novo_Cotacao1 = True Then
    USMsgBox ("Novo fornecedor cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo fornecedor"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar fornecedor"
    If CodigoLista1 <> 0 And lista_forn.ListItems.Count <> 0 Then
        lista_forn.SelectedItem = lista_forn.ListItems(CodigoLista1)
        lista_forn.SetFocus
    End If
End If
'==================================
Modulo = "Compras/Cotação"
ID_documento = txtIDListaForn
Documento = "Nº cotação: " & txtidcotacao
Documento1 = "Fornecedor: " & txtforn
ProcGravaEvento
'==================================
Novo_Cotacao1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_carteira_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)
ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_filtrar_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_un_Change()
On Error GoTo tratar_erro

If Cmb_un <> "" Then Cmb_un_com = Cmb_un

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)
If cmbfiltrarpor_necess = "Família" Then
    txtTexto_necess.Visible = False
    cmbTexto_necess.Visible = True
    ProcCarregaComboFamilia cmbTexto_necess, "familia <> 'Null' and Compras = 'True'", True
Else
    txtTexto_necess.Visible = True
    cmbTexto_necess.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)
Txtprazo_sol.Value = Date
If cmbfiltrarpor_sol = "Família" Then
    txtTexto_sol.Visible = False
    cmbTexto_sol.Visible = True
    Txtprazo_sol.Visible = False
    ProcCarregaComboFamilia cmbTexto_sol, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", True
ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
        txtTexto_sol.Visible = False
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = True
    Else
        txtTexto_sol.Visible = True
        cmbTexto_sol.Visible = False
        Txtprazo_sol.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtdesenho = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtdesenho & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcalc_peso_Click()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & Lista_itens1.SelectedItem.SubItems(3) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = False
    Compras_Cotacao = True
    Compras_Pedido = False
    Estoque_recebimento = False
    Vendas_Proposta = False
    Vendas_PI = False
    FrmCalculo_Peso.Show 1
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdconforn_Click()
On ERRO GoTo tratar_erro

Aplic = 1
Compras_Cotacao = True
Compras_Pedido = False
Vendas_Proposta = False
Vendas_PI = False
Estoque_recebimento = False
Clientes = False
Sit_REG = 0
frmCompras_pedido_DadosComerciais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcontatosforn_Click()
On ERRO GoTo tratar_erro

If txtIDforn <> "" And txtIDforn <> "0" Then
    Compras_Cotacao = True
    Compras_Pedido = False
    Financeiro_Contas_Pagar = False
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    frmCompras_Pedido_contatos.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluirForn()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With lista_forn
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) fornecedor(es)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select cotacao_fornecedor.id, cotacao_fornecedor.idforn, Cotacao_item.ID as Qtde, Cotacao_item.iditemlista, Cotacao_item.coditem, Compras_pedido_lista.ID_Requisicao, Compras_pedido_lista.ID_cotacao, Compras_pedido_lista.Status_Item FROM (cotacao_fornecedor INNER JOIN Cotacao_item ON cotacao_fornecedor.iditem = Cotacao_item.ID) INNER JOIN Compras_pedido_lista ON Cotacao_item.iditemlista = Compras_pedido_lista.IDLista where cotacao_fornecedor.idcot = " & txtidcot & " and cotacao_fornecedor.idforn = " & .ListItems(InitFor).SubItems(8), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    
                    'Excluir fornecedor do produto/serviço
                    Set TBOrdem = CreateObject("adodb.recordset")
                    TBOrdem.Open "Select compras_pedido_lista.* from compras_pedido_lista INNER JOIN compras_pedido on compras_pedido_lista.IDpedido = compras_pedido.IDpedido where compras_pedido_lista.Desenho = '" & TBFI!coditem & "' and compras_pedido.IDfornecedor = " & .ListItems(InitFor).SubItems(8) & " and compras_pedido_lista.IDlista <> " & TBFI!iditemlista, Conexao, adOpenKeyset, adLockOptimistic
                    If TBOrdem.EOF = True Then
                        Set TBCiclo = CreateObject("adodb.recordset")
                        TBCiclo.Open "Select cotacao_item.*, projproduto.codproduto from (Cotacao_item INNER JOIN Cotacao_fornecedor on cotacao_item.ID = cotacao_fornecedor.IDitem) INNER JOIN Projproduto on Projproduto.Desenho = cotacao_item.Coditem where cotacao_item.coditem = '" & TBFI!coditem & "' and cotacao_fornecedor.IDforn = " & .ListItems(InitFor).SubItems(8) & " and cotacao_item.ID <> " & TBFI!Qtde, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCiclo.EOF = True Then
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select codproduto from Projproduto where Desenho = '" & TBFI!coditem & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                Set TBAcessos = CreateObject("adodb.recordset")
                                TBAcessos.Open "Select tbl_dados_nota_fiscal.*, tbl_detalhes_nota.codproduto  from tbl_detalhes_nota INNER JOIN tbl_dados_nota_fiscal on tbl_detalhes_nota.id_nota = tbl_dados_nota_fiscal.ID where tbl_detalhes_nota.int_Cod_Produto = '" & TBFI!coditem & "' and tbl_dados_nota_fiscal.Id_Int_Cliente = " & .ListItems(InitFor).SubItems(8), Conexao, adOpenKeyset, adLockOptimistic
                                If TBAcessos.EOF = True Then Conexao.Execute "DELETE from Projproduto_fornecedor WHERE codproduto = " & TBCiclo!Codproduto & " and IDfornecedor = " & .ListItems(InitFor).SubItems(8)
                                TBAcessos.Close
                            End If
                        End If
                        TBCiclo.Close
                    End If
                    TBOrdem.Close
                    
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select cotacao_fornecedor.* FROM cotacao_fornecedor INNER JOIN Cotacao_item ON cotacao_fornecedor.iditem = Cotacao_item.ID where cotacao_fornecedor.idcot = " & txtidcot & " and Cotacao_item.iditemlista = " & TBFI!iditemlista & " and cotacao_fornecedor.idforn <> " & TBFI!IDforn, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = True Then
                        If TBFI!ID_Requisicao = 0 Then
                            Conexao.Execute "DELETE from Compras_pedido_lista where IdLista = " & TBFI!iditemlista
                            Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where IdLista = " & TBFI!iditemlista
                        Else
                            TBFI!Status_Item = "REQUISIT."
                            TBFI!ID_cotacao = 0
                            TBFI.Update
                        End If
                        Conexao.Execute "DELETE from Cotacao_item where ID = " & TBFI!Qtde
                    End If
                    TBFIltro.Close
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
            Conexao.Execute "DELETE from cotacao_fornecedor where idcot = " & txtidcot & " and IDforn = " & .ListItems(InitFor).SubItems(8)
            
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Excluir fornecedor"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Fornecedor: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) fornecedor(es) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Fornecedor(es) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposForn
    ProcCarregaListaForn False
    Lista_itens.ListItems.Clear
    Lista_itens1.ListItems.Clear
    Novo_Cotacao1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdForn_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocForn False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
Sit_REG = 2
FrmCompras_localizafornecedor.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procEmitir_PI()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
'Verificar se usuario tem acesso para aprovar cotação
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select U.IDUsuario from Usuarios U INNER JOIN Acessos A on U.IDusuario = A.IDusuario where U.IDusuario = " & pubIDUsuario & " and A.Acesso = 'Compras/Pedido'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = True Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", você não tem autorização para gerar pedido de compra."), vbExclamation, "CAPRIND v5.0"
    TBUsuarios.Close
    Exit Sub
End If
TBUsuarios.Close

If txtidcotacao.Text = "" Then
    Acao = "gerar o(s) pedido(s) de compra"
    NomeCampo = "a cotacao"
    ProcVerificaAcao
    Exit Sub
End If
ProcVerificaCamposGeraPedido

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtidcotacao.Text = "" Then
    Acao = "liberar"
    NomeCampo = "a cotacao"
    ProcVerificaAcao
    Exit Sub
End If
Liberar = True
If txtStatus <> "LIBERADA" Then ProcVerificaCamposlibera
If Liberar = True Then
    Compras_Cotacao = True
    Estoque_Requisicao = False
    frmCompras_Requisicao_aut.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovoForn()
On ERRO GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If ProcVerifSatus("criar um novo fornecedor", True) = False Then Exit Sub
ProcLimpaCamposForn
Frame8.Enabled = True
Novo_Cotacao1 = True

ProcConfVariaveisLocForn False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
Sit_REG = 2
FrmCompras_localizafornecedor.Show 1
If txtforn <> "" Then cmdcontatosforn.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaForn(CheckBox As Boolean)
On Error GoTo tratar_erro

Fornecedor = ""
lista_forn.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from cotacao_fornecedor where idcot = " & Cont & " order by forn", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        If Fornecedor <> TBLISTA!forn Then
            With lista_forn.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!forn), "", TBLISTA!forn)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!contforn), "", TBLISTA!contforn)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!telforn), "", TBLISTA!telforn)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!faxforn), "", TBLISTA!faxforn)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!CondPagtoforn), "", TBLISTA!CondPagtoforn)
                If TBLISTA!CIFforn = True Then .Item(.Count).SubItems(6) = "SIM" Else .Item(.Count).SubItems(6) = "NÃO"
                If TBLISTA!FOBforn = True Then .Item(.Count).SubItems(7) = "SIM" Else .Item(.Count).SubItems(7) = "NÃO"
                .Item(.Count).SubItems(8) = TBLISTA!IDforn
                
                If CheckBox = True Then .Item(.Count).Checked = True
            End With
        End If
        Fornecedor = TBLISTA!forn
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtidcot = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_Cotacao order by Id_cotacao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Id_cotacao = " & txtidcot)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtidcot = TBLISTA!ID_cotacao
        Set TBCotacao = CreateObject("adodb.recordset")
        TBCotacao.Open "Select * from Compras_Cotacao where Id_cotacao = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCamposCotacao
        ProcLimpaCamposItem
        ProcLimpaCamposForn
        Lista_itens1.ListItems.Clear
        lista_forn.ListItems.Clear
        ProcCarregaDados
        ProcCarregaListaForn False
    Else
        USMsgBox ("Fim dos cadastros de cotação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Cotacao2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procProximoForn()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Forn from Cotacao_fornecedor where idcot = " & txtidcot & " Group by forn", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    TBLISTA.Find ("Forn = '" & txtforn & "'")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Cotacao_fornecedor where idcot = " & txtidcot & " and Forn = '" & TBLISTA!forn & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcLimpaCamposForn
            ProcLimpaCamposItem
            ProcCarregaDadosForn
            ProcCarregaListaItens1
        End If
    Else
        USMsgBox ("Fim dos cadastros de fornecedores."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Cotacao2 = False

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
If ProcVerifSatus("alterar o produto/serviço", True) = False Then Exit Sub
Acao = "salvar"
If txtdesenho = "" Then
    NomeCampo = "o produto/serviço"
    ProcVerificaAcao
    Exit Sub
End If
If Cmb_familia.Text = "" Then
    NomeCampo = "a família"
    ProcVerificaAcao
    Cmb_familia.SetFocus
    Exit Sub
End If
If Cmb_un.Text = "" Then
    NomeCampo = "a unidade de estoque"
    ProcVerificaAcao
    Cmb_un.SetFocus
    Exit Sub
End If
If Cmb_un_com.Text = "" Then
    NomeCampo = "a unidade comercial"
    ProcVerificaAcao
    Cmb_un_com.SetFocus
    Exit Sub
End If
valor = IIf(Txt_quantidade = "", 0, Txt_quantidade)
If valor <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    Txt_quantidade.SetFocus
    Exit Sub
End If
If txtDescricao_item = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricao_item.SetFocus
    Exit Sub
End If
If txtDescricao_comercial = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    txtDescricao_comercial.SetFocus
    Exit Sub
End If
If txtPrazoentregaforn <> "__/__/____" Then
    If IsDate(txtPrazoentregaforn) = False Then
        NomeCampo = "o prazo de entrega"
        ProcVerificaAcao
        txtPrazoentregaforn.SetFocus
        Exit Sub
    End If
End If

'If cmbreferencia <> "" Then If FunVerifiCodRefUtilizado(txtDesenho, cmbreferencia) = True Then Exit Sub    'Verifica se o código de referencia está vinculado a outro produto

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from compras_pedido_lista where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!ID_cotacao = Cont
    TBGravar!Descricao = txtDescricao_item
    TBGravar!Descricao_comercial = txtDescricao_comercial
    TBGravar!quant_req = Txt_quantidade.Text
    TBGravar!quant_req_PC = IIf(Txt_quantidade_PC = "", Null, Txt_quantidade_PC)
    TBGravar!Un = Cmb_un
    TBGravar!Unidade_com = Cmb_un_com
    TBGravar!Familia = Cmb_familia
    TBGravar!Obs_cotacao = txtObs
    TBGravar.Update
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Cotacao_fornecedor where ID = " & Txt_ID_tabela_forn, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!IDitem = txtiditem
    If cmbReferencia <> "" Then
        TBGravar!N_referencia = cmbReferencia
        
        'Grava código de referência no produto
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Codproduto from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from item_aplicacoes where Codproduto = " & TBAbrir!Codproduto & " and n_referencia = '" & cmbReferencia & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = True Then TBProduto.AddNew
            TBProduto!Codproduto = TBAbrir!Codproduto
            TBProduto!N_referencia = cmbReferencia
            TBProduto!Descricao = IIf(txtDescricao_item.Text = "", Null, txtDescricao_item.Text)
            TBProduto!Aplicacao = txtforn1
            TBProduto!ID_cliente_forn = txtidforn1
            TBProduto!Tipo = "F"
            TBProduto.Update
            TBProduto.Close
        End If
        TBAbrir.Close
    Else
        TBGravar!N_referencia = Null
    End If
    TBGravar!prazoentregaforn = IIf(txtPrazoentregaforn = "__/__/____", Null, txtPrazoentregaforn)
    TBGravar!precounit = IIf(txtvalorunitforn = "", Null, txtvalorunitforn)
    TBGravar!Desconto = IIf(txtDesconto = "", Null, txtDesconto)
    TBGravar!ValorDesconto = IIf(txtvalordesconto = "", Null, txtvalordesconto)
    TBGravar!preco_unitario_desconto = IIf(txtvalorunitariodesc = "", Null, txtvalorunitariodesc)
    TBGravar!IPI = IIf(txtIPI = "", Null, txtIPI)
    TBGravar!VlrIPI = IIf(TxtvlrIpi = "", Null, TxtvlrIpi)
    TBGravar!ICMS = IIf(txtICMS = "", Null, txtICMS)
    TBGravar!vlrICMS = IIf(txtvlrICMS = "", Null, txtvlrICMS)
    TBGravar!PrecoTotal = IIf(txttotalforn = "", Null, txttotalforn)
    TBGravar!Obsforn = Txt_obs_fornecedor
    TBGravar.Update
End If
TBGravar.Close
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Compras/Cotação"
Evento = "Alterar produto/serviço"
ID_documento = TXTIDLista
Documento = "Nº cotação: " & txtidcotacao
Documento1 = "Fornecedor: " & txtforn & " - Cód. interno: " & txtdesenho
ProcGravaEvento
'==================================
ProcCarregaListaItens1
If CodigoLista2 <> 0 And Lista_itens1.ListItems.Count <> 0 Then
    Lista_itens1.SelectedItem = Lista_itens1.ListItems(CodigoLista2)
    Lista_itens1.SetFocus
End If
ProcCarregaListaItens

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab2.Tab = 0 Then
            If TBLISTA_Cotacao_Necessidade.AbsolutePage <> 2 Then
                If TBLISTA_Cotacao_Necessidade.AbsolutePage = -3 Then
                    ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.PageCount - 1)
                Else
                    TBLISTA_Cotacao_Necessidade.AbsolutePage = TBLISTA_Cotacao_Necessidade.AbsolutePage - 2
                    ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.AbsolutePage)
                End If
            Else
                ProcExibePagina_Necessidade (1)
            End If
        Else
            If TBLISTA_Cotacao_Solicitacao.AbsolutePage <> 2 Then
                If TBLISTA_Cotacao_Solicitacao.AbsolutePage = -3 Then
                    ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.PageCount - 1)
                Else
                    TBLISTA_Cotacao_Solicitacao.AbsolutePage = TBLISTA_Cotacao_Solicitacao.AbsolutePage - 2
                    ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.AbsolutePage)
                End If
            Else
                ProcExibePagina_Solicitacao (1)
            End If
        End If
    Case 1:
        If TBLISTA_Compras_Cotacao.AbsolutePage <> 2 Then
            If TBLISTA_Compras_Cotacao.AbsolutePage = -3 Then
                ProcExibePagina (TBLISTA_Compras_Cotacao.PageCount - 1)
            Else
                TBLISTA_Compras_Cotacao.AbsolutePage = TBLISTA_Compras_Cotacao.AbsolutePage - 2
                ProcExibePagina (TBLISTA_Compras_Cotacao.AbsolutePage)
            End If
        Else
            ProcExibePagina (1)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4))
If Quant <= 1 Or txtPagIr(index) > Quant Then Exit Sub
If txtPagIr(index).Text >= 1 And txtPagIr(index).Text <= Quant Then
    Select Case SSTab1.Tab
        Case 0:
            If SSTab2.Tab = 0 Then
                TBLISTA_Cotacao_Necessidade.AbsolutePage = txtPagIr(index).Text
                ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.AbsolutePage)
            Else
                TBLISTA_Cotacao_Solicitacao.AbsolutePage = txtPagIr(index).Text
                ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.AbsolutePage)
            End If
        Case 1:
            TBLISTA_Compras_Cotacao.AbsolutePage = txtPagIr(index).Text
            ProcExibePagina (TBLISTA_Compras_Cotacao.AbsolutePage)
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab2.Tab = 0 Then
            TBLISTA_Cotacao_Necessidade.AbsolutePage = 1
            ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.AbsolutePage)
        Else
            TBLISTA_Cotacao_Solicitacao.AbsolutePage = 1
            ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.AbsolutePage)
        End If
    Case 1:
        TBLISTA_Compras_Cotacao.AbsolutePage = 1
        ProcExibePagina (TBLISTA_Compras_Cotacao.AbsolutePage)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab2.Tab = 0 Then
            If TBLISTA_Cotacao_Necessidade.AbsolutePage <> -3 Then
                If TBLISTA_Cotacao_Necessidade.AbsolutePage = 1 Then
                    ProcExibePagina_Necessidade (2)
                Else
                    ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.AbsolutePage)
                End If
            Else
                ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.PageCount)
            End If
        Else
            If TBLISTA_Cotacao_Solicitacao.AbsolutePage <> -3 Then
                If TBLISTA_Cotacao_Solicitacao.AbsolutePage = 1 Then
                    ProcExibePagina_Solicitacao (2)
                Else
                    ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.AbsolutePage)
                End If
            Else
                ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.PageCount)
            End If
        End If
    Case 1:
        If TBLISTA_Compras_Cotacao.AbsolutePage <> -3 Then
            If TBLISTA_Compras_Cotacao.AbsolutePage = 1 Then
                ProcExibePagina (2)
            Else
                ProcExibePagina (TBLISTA_Compras_Cotacao.AbsolutePage)
            End If
        Else
            ProcExibePagina (TBLISTA_Compras_Cotacao.PageCount)
        End If
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click(index As Integer)
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas(index).Caption, 4)) <= 1 Then Exit Sub
Select Case SSTab1.Tab
    Case 0:
        If SSTab2.Tab = 0 Then
            TBLISTA_Cotacao_Necessidade.AbsolutePage = TBLISTA_Cotacao_Necessidade.PageCount
            ProcExibePagina_Necessidade (TBLISTA_Cotacao_Necessidade.AbsolutePage)
        Else
            TBLISTA_Cotacao_Solicitacao.AbsolutePage = TBLISTA_Cotacao_Solicitacao.PageCount
            ProcExibePagina_Solicitacao (TBLISTA_Cotacao_Solicitacao.AbsolutePage)
        End If
    Case 1:
        TBLISTA_Compras_Cotacao.AbsolutePage = TBLISTA_Compras_Cotacao.PageCount
        ProcExibePagina (TBLISTA_Compras_Cotacao.AbsolutePage)
End Select

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
            Case vbKeyF2: If SSTab2.Tab = 0 Then ProcFiltrar_Necessidade Else ProcFiltrar_Solicitacao
            Case vbKeyF3: ProcGerarCot
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF8: ProcLiberar
            Case vbKeyF10: procEmitir_PI
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: procNovoForn
            Case vbKeyF3: procSalvarForn
            Case vbKeyF4: procExcluirForn
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: procAdd_item
            Case vbKeyF8: procAprovarForn
            Case vbKeyF9: procNao_aprovar
            Case vbkey10: procCancelar_Aprovacao
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyF3: procSalvar_item
            Case vbKeyF4: procExcluir_item
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: procExcluirTodosForn
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

ProcCarregaToolBar1 Me, 15105, 5, True
ProcCarregaToolBar3 Me, 15195, 14, True
ProcCarregaToolBar4 Me, 15195, 14, True
ProcCarregaToolBar5 Me, 15105, 10, True
ProcCarregaToolBar6 Me, 15105, 6, True
Formulario = "Compras/Cotação"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
SSTab2.Tab = 0
SSTab3.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboEmpresa Cmb_empresa_carteira, False
ProcCarregaComboFamilia Cmb_familia, "familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", False
ProcCarregaComboUnidade Cmb_un, False
ProcCarregaComboUnidade Cmb_un_com, False

ProcFiltroPadrao cmbfiltrarpor_necess, Optmeio_necess, Optfim_necess, optIgual_necess, Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex), "Produtos/Serviços", "C", True
If Permitido = False Then cmbfiltrarpor_necess = "Código interno"
ProcFiltroPadrao cmbfiltrarpor_sol, optMeio_sol, optFim_sol, optIgual_sol, Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex), "Produtos/Serviços", "C", True
If Permitido = False Then cmbfiltrarpor_sol = "Código interno"
Cmb_filtrar = "Com necessidade"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Cotação"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362C" Then frmCompras_reqcot_atualizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmCompras_reqcot_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza número das cotações
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select id_cotacao, dataemissao, Cotacaotexto from Compras_Cotacao order by id_cotacao", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                TBCompras.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCompras.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCompras.MoveFirst
                Do While TBCompras.EOF = False
                    Cont = TBCompras!ID_cotacao
                    a = Cont
                    Ano = Right(Year(TBCompras!DataEmissao), 2)
                    Select Case Len(a)
                        Case 1: a = "COT-0000" & Cont & "/" & Ano
                        Case 2: a = "COT-000" & Cont & "/" & Ano
                        Case 3: a = "COT-00" & Cont & "/" & Ano
                        Case 4: a = "COT-0" & Cont & "/" & Ano
                        Case 5: a = "COT-" & Cont & "/" & Ano
                    End Select
                    TBCompras!Cotacaotexto = a
                    TBCompras.Update
                    TBCompras.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
            
        If .Chk2.Value = 1 Then
            'Atualiza desconto dos produtos
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from Cotacao_fornecedor", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                TBCompras.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCompras.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCompras.MoveFirst
                Do While TBCompras.EOF = False
                    If IsNull(TBCompras!Desconto) = True Or TBCompras!Desconto = "" Then TBCompras!Desconto = 0
                    If IsNull(TBCompras!ValorDesconto) = True Or TBCompras!ValorDesconto = "" Then TBCompras!ValorDesconto = 0
                    If IsNull(TBCompras!preco_unitario_desconto) = True Or TBCompras!preco_unitario_desconto = "" Then TBCompras!preco_unitario_desconto = TBCompras!precounit
                    TBCompras.Update
                    TBCompras.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBCompras.Close
        End If
        
        If .Chk3.Value = 1 Then
            'Atualiza produtos sem fornecedores
            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from Cotacao_fornecedor where IDitem <> 0 order by IDitem", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                TBCompras_Pedido.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCompras_Pedido.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCompras_Pedido.MoveFirst
                Do While TBCompras_Pedido.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Cotacao_item where ID = " & TBCompras_Pedido!IDitem, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        Conexao.Execute "DELETE from Cotacao_fornecedor where ID = " & TBCompras_Pedido!ID
                    End If
                    TBAbrir.Close
                    TBCompras_Pedido.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBCompras_Pedido.Close
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Compras/Cotação"
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

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtidcotacao.Text = "" Then
    Acao = "visualizar impressão"
    NomeCampo = "a cotacao"
    ProcVerificaAcao
    Exit Sub
End If
frmcompras_reqcot_menu_impressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Img_calendario_validade_Click()
On Error GoTo tratar_erro

ProcAbrirCalendario 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

If txtPrazoentregaforn.Enabled = False Then Exit Sub
ProcAbrirCalendario 2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirCalendario(Opcao As Integer)
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = False
Compras_Fallow_up = False
Vendas_Carteira = False
Vendas_Proposta = False
Vendas_PI = False
Manutencao = False
Compras_Cotacao = True
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
Compras_Fornecedores = False
Vendas_Programacao = False
Outros_solicitacaoPCP = False
Estoque_recebimento = False
Sit_Data = Opcao
FrmCalendario.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Necessidade()
On Error GoTo tratar_erro

If Opt_PCP.Value = True Then NomeTabela = "Cotacao_Necessidade" Else NomeTabela = "Cotacao_Necessidade_PIEST"
CamposFiltro = "CN.Codproduto, CN.Desenho, CN.Descricao, CN.Unidade, CN.Unidade_Com, CN.Necessidade, CN.Necessidade_estoque, CN.Necessidade_PC"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (" & NomeTabela & " CN LEFT JOIN item_aplicacoes IA ON CN.codproduto = IA.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CN.codproduto"
If Cmb_filtrar = "Com necessidade" Then TextoFiltroEstoque = " and CN.Necessidade > 0" Else TextoFiltroEstoque = " and CN.Necessidade_estoque > 0"
TextoFiltroPadrao = "CN.ID_empresa = " & Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex) & TextoFiltroEstoque & " group by " & CamposFiltro & " order by CN.desenho"

If txtTexto_necess.Visible = True And txtTexto_necess <> "" Or cmbTexto_necess.Visible = True And cmbTexto_necess <> "" Then
    If cmbfiltrarpor_necess = "Família" Then
        StrSql_Cotacao_Necessidade = INNERJOINTEXTO & " where CN.classe = '" & cmbTexto_necess & "' and " & TextoFiltroPadrao
    Else
        Select Case cmbfiltrarpor_necess
            Case "Código interno": TextoFiltro = "CN.Desenho"
            Case "Código de referência": TextoFiltro = "IA.n_referencia"
            Case "Descrição": TextoFiltro = "CN.Descricao"
            Case "Part number": TextoFiltro = "PFAB.Part_number"
        End Select
        StrSql_Cotacao_Necessidade = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_necess, Optmeio_necess, Optfim_necess, optIgual_necess, txtTexto_necess) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Cotacao_Necessidade = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Necessidade

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Necessidade()
On Error GoTo tratar_erro

If StrSql_Cotacao_Necessidade = "" Then Exit Sub
lblRegistros(0).Caption = "Nº de registros: 0"
lblPaginas(0).Caption = "Página: 0 de: 0"
ListaNecessidade.ListItems.Clear
Set TBLISTA_Cotacao_Necessidade = CreateObject("adodb.recordset")
TBLISTA_Cotacao_Necessidade.Open StrSql_Cotacao_Necessidade, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Cotacao_Necessidade.EOF = False Then ProcExibePagina_Necessidade (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Necessidade(Pagina)
On Error GoTo tratar_erro

ListaNecessidade.ListItems.Clear
TBLISTA_Cotacao_Necessidade.PageSize = IIf(txtNreg(0) = "", 30, txtNreg(0))
TBLISTA_Cotacao_Necessidade.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Cotacao_Necessidade.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Cotacao_Necessidade.RecordCount - IIf(Pagina > 1, (TBLISTA_Cotacao_Necessidade.PageSize * (Pagina - 1)), 0), TBLISTA_Cotacao_Necessidade.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Cotacao_Necessidade.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaNecessidade.ListItems.Add(, , TBLISTA_Cotacao_Necessidade!Codproduto)
        .SubItems(1) = IIf(IsNull(TBLISTA_Cotacao_Necessidade!Desenho), "", TBLISTA_Cotacao_Necessidade!Desenho)
        .SubItems(2) = IIf(IsNull(TBLISTA_Cotacao_Necessidade!Descricao), "", TBLISTA_Cotacao_Necessidade!Descricao)
        .SubItems(3) = IIf(IsNull(TBLISTA_Cotacao_Necessidade!Unidade_com), "", TBLISTA_Cotacao_Necessidade!Unidade_com)
        If TBLISTA_Cotacao_Necessidade!Unidade <> TBLISTA_Cotacao_Necessidade!Unidade_com Then
            If Cmb_filtrar = "Com necessidade" Then qt = Format(TBLISTA_Cotacao_Necessidade!Necessidade, "###,##0.0000") Else qt = Format(TBLISTA_Cotacao_Necessidade!Necessidade_estoque, "###,##0.0000")
            If FunVerifUNConversao(TBLISTA_Cotacao_Necessidade!Unidade, TBLISTA_Cotacao_Necessidade!Unidade_com) = True Then
                Qtde = FunConverteUN(TBLISTA_Cotacao_Necessidade!Unidade_com, TBLISTA_Cotacao_Necessidade!Unidade, qt, TBLISTA_Cotacao_Necessidade!Desenho)
                .SubItems(4) = Format(Qtde, "###,##0.0000")
            Else
                If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLISTA_Cotacao_Necessidade!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLISTA_Cotacao_Necessidade!Necessidade_estoque, "###,##0.0000")
            End If
        Else
            If Cmb_filtrar = "Com necessidade" Then .SubItems(4) = Format(TBLISTA_Cotacao_Necessidade!Necessidade, "###,##0.0000") Else .SubItems(4) = Format(TBLISTA_Cotacao_Necessidade!Necessidade_estoque, "###,##0.0000")
        End If
        .SubItems(5) = TBLISTA_Cotacao_Necessidade!Necessidade_PC
        If Cmb_filtrar = "Com necess. estoque" Then NReal = Format(TBLISTA_Cotacao_Necessidade!Necessidade_estoque, "###,##0.0000") Else NReal = Format(TBLISTA_Cotacao_Necessidade!Necessidade, "###,##0.0000")
        If NReal > 0 Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .ListSubItems(3).ForeColor = vbRed
            .ListSubItems(4).ForeColor = vbRed
            .ListSubItems(5).ForeColor = vbRed
        End If
    End With
    TBLISTA_Cotacao_Necessidade.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(0).Caption = "Nº de registros: " & TBLISTA_Cotacao_Necessidade.RecordCount
If TBLISTA_Cotacao_Necessidade.AbsolutePage = adPosBOF Then
   lblPaginas(0).Caption = "Página: 1 de: " & TBLISTA_Cotacao_Necessidade.PageCount
ElseIf TBLISTA_Cotacao_Necessidade.AbsolutePage = adPosEOF Then
        lblPaginas(0).Caption = "Página: " & TBLISTA_Cotacao_Necessidade.PageCount & " de: " & TBLISTA_Cotacao_Necessidade.PageCount
    Else
        lblPaginas(0).Caption = "Página: " & TBLISTA_Cotacao_Necessidade.AbsolutePage - 1 & " de: " & TBLISTA_Cotacao_Necessidade.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Solicitacao()
On Error GoTo tratar_erro

CamposFiltro = "CR.ID_requisicao, CR.Requisicaotexto, CPL.IdLista, CPL.Status_Item, CPL.desenho, CPL.descricao, CPL.Un, CPL.Unidade_com, CPL.quant_req, CPL.quant_req_PC, CPL.detalheitem, CPL.prazoreq, CPL.obs"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (Compras_requisicao CR INNER JOIN Compras_pedido_lista CPL ON CR.ID_Requisicao = CPL.ID_Requisicao) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = CPL.codproduto"
TextoFiltroPadrao = "CPL.status_item = 'REQUISIT.' and CR.Status = 'LIBERADA' group by " & CamposFiltro & " order by CR.ID_requisicao"

If txtTexto_sol.Visible = True And txtTexto_sol <> "" Or cmbTexto_sol.Visible = True And cmbTexto_sol <> "" Or Txtprazo_sol.Visible = True Then
    If cmbfiltrarpor_sol = "Família" Then
        StrSql_Cotacao_Solicitacao = INNERJOINTEXTO & " where CPL.Familia = '" & cmbTexto_sol & "' and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor_sol = "Prazo entrega" Then
            StrSql_Cotacao_Solicitacao = INNERJOINTEXTO & " where CPL.Prazoreq = '" & Format(Txtprazo_sol.Value, "Short Date") & "' and " & TextoFiltroPadrao
        Else
            Select Case cmbfiltrarpor_sol
                Case "Solicitação": TextoFiltro = "CR.Requisicaotexto"
                Case "Código interno": TextoFiltro = "CPL.desenho"
                Case "Descrição": TextoFiltro = "CPL.descricao"
                Case "Descrição comercial": TextoFiltro = "CPL.descricao_comercial"
                Case "Detalhe": TextoFiltro = "CPL.Detalheitem"
                Case "Part number": TextoFiltro = "PFAB.Part_number"
            End Select
            StrSql_Cotacao_Solicitacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(optInicio_sol, optMeio_sol, optFim_sol, optIgual_sol, txtTexto_sol) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Cotacao_Solicitacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcCarregalista_Solicitacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_Solicitacao()
On Error GoTo tratar_erro

lblRegistros(1).Caption = "Nº de reg.: 0"
lblPaginas(1).Caption = "Página: 0 de: 0"
Lista_solicitados.ListItems.Clear
If StrSql_Cotacao_Solicitacao = "" Then Exit Sub
Set TBLISTA_Cotacao_Solicitacao = CreateObject("adodb.recordset")
TBLISTA_Cotacao_Solicitacao.Open StrSql_Cotacao_Solicitacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Cotacao_Solicitacao.EOF = False Then ProcExibePagina_Solicitacao (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina_Solicitacao(Pagina)
On Error GoTo tratar_erro

Lista_solicitados.ListItems.Clear
TBLISTA_Cotacao_Solicitacao.PageSize = IIf(txtNreg(1) = "", 30, txtNreg(1))
TBLISTA_Cotacao_Solicitacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Cotacao_Solicitacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Cotacao_Solicitacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Cotacao_Solicitacao.PageSize * (Pagina - 1)), 0), TBLISTA_Cotacao_Solicitacao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Cotacao_Solicitacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_solicitados.ListItems
        .Add , , TBLISTA_Cotacao_Solicitacao!IDlista
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Status_Item), "", TBLISTA_Cotacao_Solicitacao!Status_Item)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Requisicaotexto), "", TBLISTA_Cotacao_Solicitacao!Requisicaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Desenho), "", TBLISTA_Cotacao_Solicitacao!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Descricao), "", TBLISTA_Cotacao_Solicitacao!Descricao)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Un), "", TBLISTA_Cotacao_Solicitacao!Un)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Unidade_com), "", TBLISTA_Cotacao_Solicitacao!Unidade_com)
        If TBLISTA_Cotacao_Solicitacao!Un <> TBLISTA_Cotacao_Solicitacao!Unidade_com Then valor = FunConversaoFinalUn(TBLISTA_Cotacao_Solicitacao!Un, TBLISTA_Cotacao_Solicitacao!Unidade_com, TBLISTA_Cotacao_Solicitacao!quant_req, TBLISTA_Cotacao_Solicitacao!Desenho, True) Else valor = TBLISTA_Cotacao_Solicitacao!quant_req
        .Item(.Count).SubItems(7) = FunFormataCasasDecimais(4, valor)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!quant_req), "", FunFormataCasasDecimais(4, TBLISTA_Cotacao_Solicitacao!quant_req))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!detalheitem), "", TBLISTA_Cotacao_Solicitacao!detalheitem)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!prazoreq), "", Format(TBLISTA_Cotacao_Solicitacao!prazoreq, "dd/mm/yy"))
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Cotacao_Solicitacao!Obs), "", TBLISTA_Cotacao_Solicitacao!Obs)
    End With
    ContadorReg = ContadorReg + 1
    TBLISTA_Cotacao_Solicitacao.MoveNext
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(1).Caption = "Nº de reg.: " & TBLISTA_Cotacao_Solicitacao.RecordCount
If TBLISTA_Cotacao_Solicitacao.AbsolutePage = adPosBOF Then
   lblPaginas(1).Caption = "Pág.: 1 de: " & TBLISTA_Cotacao_Solicitacao.PageCount
ElseIf TBLISTA_Cotacao_Solicitacao.AbsolutePage = adPosEOF Then
        lblPaginas(1).Caption = "Pág.: " & TBLISTA_Cotacao_Solicitacao.PageCount & " de: " & TBLISTA_Cotacao_Solicitacao.PageCount
    Else
        lblPaginas(1).Caption = "Pág.: " & TBLISTA_Cotacao_Solicitacao.AbsolutePage - 1 & " de: " & TBLISTA_Cotacao_Solicitacao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmCompras_reqcot_abrir_cotadas.Show 1

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
If txtidcotacao.Text = "" Then
    Acao = "alterar o status"
    NomeCampo = "a cotacao"
    ProcVerificaAcao
    Exit Sub
End If
If txtStatus = "APROVADA" Or txtStatus = "LIBERADA" Then
    USMsgBox ("Não é permitido alterar o status, pois o status da mesma está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select IDlista from Compras_pedido_lista where id_cotacao = " & txtidcot & " and Status_Item <> 'COTANDO' and Status_Item <> 'REQUISIT.' and Status_Item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido alterar o status da cotação, pois os produtos/serviços já sofreram alguma alteração."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
If txtStatus = "CANCELADA" Then
    If USMsgBox("Deseja realmente alterar o status da cotação " & txtidcotacao & " para cotando?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'COTANDO' where id_cotacao = " & txtidcot
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Compras_Cotacao where Cotacaotexto = '" & txtidcotacao & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBAbrir!statuscotacao = "COTANDO"
            TBAbrir!deletou = ""
            TBAbrir!datacancelada = Null
            TBAbrir!motivo = ""
            TBAbrir.Update
            TBAbrir.Close
            USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Alterar status da cotação"
            ID_documento = Cont
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = ""
            ProcGravaEvento
            '==================================
            ProcCarregaListaCotacao (IIf(ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5))))
            If CodigoLista <> 0 And lista_cot.ListItems.Count <> 0 Then
                lista_cot.SelectedItem = lista_cot.ListItems(CodigoLista)
                lista_cot.SetFocus
            End If
            Set TBCotacao = CreateObject("adodb.recordset")
            TBCotacao.Open "Select * from compras_cotacao where cotacaotexto = '" & txtidcotacao & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCotacao.EOF = False Then
                ProcLimpaCamposCotacao
                ProcCarregaDados
            End If
            TBCotacao.Close
        End If
    Else
        frmCompras_reqcot_cancelar.Show 1
    End If
Else
    frmCompras_reqcot_cancelar.Show 1
    If txtStatus = "CANCELADA" Then
        ProcCarregaListaCotacao (IIf(ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5))))
        If CodigoLista <> 0 And lista_cot.ListItems.Count <> 0 Then
            lista_cot.SelectedItem = lista_cot.ListItems(CodigoLista)
            lista_cot.SetFocus
        End If
        Set TBCotacao = CreateObject("adodb.recordset")
        TBCotacao.Open "Select * from compras_cotacao where cotacaotexto = '" & txtidcotacao & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBCotacao.EOF = False Then
            ProcLimpaCamposCotacao
            ProcCarregaDados
        End If
        TBCotacao.Close
    End If
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
With Lista_itens1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select cotacao_fornecedor.id, cotacao_fornecedor.idforn, cotacao_fornecedor.iditem, Cotacao_item.id, Cotacao_item.iditemlista, Compras_pedido_lista.ID_Requisicao, Compras_pedido_lista.ID_cotacao, Compras_pedido_lista.Status_Item FROM (cotacao_fornecedor INNER JOIN Cotacao_item ON cotacao_fornecedor.iditem = Cotacao_item.ID) INNER JOIN Compras_pedido_lista ON Cotacao_item.iditemlista = Compras_pedido_lista.IDLista where cotacao_fornecedor.idcot = " & txtidcot & " and cotacao_fornecedor.idforn = " & txtIDforn & " and Cotacao_item.ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                
                'Verifica se tem outro produto para este fornecedor
                Permitido1 = False
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from cotacao_fornecedor where idcot = " & txtidcot & " and idforn = " & TBFI!IDforn & " and IDitem <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    Permitido1 = True
                Else
                    TBFI!IDitem = 0
                    TBFI.Update
                End If
                
                'Verifica se o produto esta vinculado a outro fornecedor
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select cotacao_fornecedor.* FROM cotacao_fornecedor INNER JOIN Cotacao_item ON cotacao_fornecedor.iditem = Cotacao_item.ID where cotacao_fornecedor.idcot = " & txtidcot & " and Cotacao_item.iditemlista = " & TBFI!iditemlista & " and cotacao_fornecedor.idforn <> " & TBFI!IDforn & " and Cotacao_item.ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = True Then
                    If TBFI!ID_Requisicao = 0 Then
                        Conexao.Execute "DELETE from Compras_pedido_lista where IdLista = " & TBFI!iditemlista
                        Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where IdLista = " & TBFI!iditemlista
                    Else
                        TBFI!Status_Item = "REQUISIT."
                        TBFI!ID_cotacao = 0
                        TBFI.Update
                    End If
                    Conexao.Execute "DELETE from Cotacao_item where ID = " & TBFI!ID
                End If
                TBFIltro.Close
                
                If Permitido1 = True Then Conexao.Execute "DELETE from cotacao_fornecedor where idcot = " & txtidcot & " and idforn = " & TBFI!IDforn & " and IDitem = " & .ListItems(InitFor)
                
            End If
            TBFI.Close
            
            'Excluir fornecedor do produto/serviço
            Set TBOrdem = CreateObject("adodb.recordset")
            TBOrdem.Open "Select Compras_pedido_lista.* from Compras_pedido_lista INNER JOIN Compras_pedido on Compras_pedido_lista.IDpedido = Compras_pedido.IDpedido where Compras_pedido_lista.desenho = '" & .ListItems(InitFor).SubItems(3) & "' and Compras_pedido.idfornecedor = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = True Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select cotacao_item.* from Cotacao_item INNER JOIN Cotacao_fornecedor on cotacao_item.ID = cotacao_fornecedor.IDitem where cotacao_item.coditem = '" & .ListItems(InitFor).SubItems(3) & "' and cotacao_fornecedor.idforn = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = True Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select tbl_dados_nota_fiscal.*, tbl_detalhes_nota.codproduto  from tbl_detalhes_nota INNER JOIN tbl_dados_nota_fiscal on tbl_detalhes_nota.id_nota = tbl_dados_nota_fiscal.ID where tbl_detalhes_nota.int_Cod_Produto = '" & .ListItems(InitFor).SubItems(3) & "' and tbl_dados_nota_fiscal.Id_Int_Cliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then Conexao.Execute "DELETE from Projproduto_fornecedor WHERE codproduto = " & .ListItems(InitFor) & " and IDfornecedor = " & txtIDforn
                    TBFI.Close
                End If
                TBItem.Close
            End If
            TBOrdem.Close
            
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Excluir produto/serviço do fornecedor"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Fornecedor: " & txtforn & " - Cód. interno: " & .ListItems(InitFor).SubItems(3)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("produto(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposItem
    ProcCarregaListaItens
    ProcCarregaListaItens1
    Frame4.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAdd_item()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If ProcVerifSatus("adicionar o(s) produto(s)/serviço(s)", True) = False Then Exit Sub
With lista_forn
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then GoTo 1
    Next InitFor
End With
USMsgBox ("Selecione o(s) fornecedor(es) na lista antes de adicionar o(s) produto(s)/serviço(s) nesta cotação."), vbExclamation, "CAPRIND v5.0"
Exit Sub
1:
    frmCompras_reqcot_abrir.Show 1

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
If Cmb_empresa = "" Then
    USMsgBox ("Informe a empresa antes de criar uma nova cotação."), vbExclamation, "CAPRIND v5.0"
    Cmb_empresa.SetFocus
    Exit Sub
End If
ProcLimpaCamposCotacao
ProcLimparTudo
Frame2.Enabled = True
txtObservacao.SetFocus
Novo_Cotacao = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame8.Enabled = False
Frame9.Enabled = False
Frame4.Enabled = False
ProcLimpaCamposForn
lista_forn.ListItems.Clear
Lista_itens.ListItems.Clear
ProcLimpaCamposItem
Lista_itens1.ListItems.Clear
Novo_Cotacao1 = False
Novo_Cotacao2 = False
SSTab3.Tab = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Cotacao1 = True Then
    If USMsgBox("O fornecedor ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvarForn
        If Novo_Cotacao1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Cotacao2 = True Then
    If USMsgBox("O produto/serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar_item
        If Novo_Cotacao2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Cotacao1 = False
Novo_Cotacao2 = False
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
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If ProcVerifSatus("alterar a cotação", True) = False Then Exit Sub
If Txt_data_validade <> "__/__/____" Then
    If IsDate(Txt_data_validade) = False Then
        USMsgBox ("A data de validade foi digitada incorretamente."), vbExclamation, "CAPRIND v5.0"
        Txt_data_validade.SetFocus
        Exit Sub
    End If
End If
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * from compras_cotacao where id_cotacao = " & IIf(txtidcot = "", 0, txtidcot), Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = True Then
    TBCotacao.AddNew
    TBCotacao!statuscotacao = "COTANDO"
    TBCotacao!Cotacaotexto = ProcCriarNovoNumero
End If
TBCotacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBCotacao!DataEmissao = IIf(txtDataemissao = "", Date, txtDataemissao)
TBCotacao!Requisitante = IIf(txtRequisitante = "", pubUsuario, txtRequisitante)
TBCotacao!Setor = IIf(cmbSetor = "", pubSetor, cmbSetor)
TBCotacao!Data_validade = IIf(Txt_data_validade = "__/__/____", Null, Txt_data_validade)
TBCotacao!Obs = txtObservacao
TBCotacao.Update
ProcCarregaDados
TBCotacao.Close

If Novo_Cotacao = True Then
    USMsgBox ("Nova cotação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Cotacao_Localizar = "Select * from Compras_Cotacao where Cotacaotexto = '" & txtidcotacao & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    ProcCarregaListaCotacao (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaListaCotacao (IIf(ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5))))
    If CodigoLista <> 0 And lista_cot.ListItems.Count <> 0 Then
        lista_cot.SelectedItem = lista_cot.ListItems(CodigoLista)
        lista_cot.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Compras/Cotação"
    ID_documento = Cont
    Documento = "Nº cotação: " & txtidcotacao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_Cotacao = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposCotacao()
On Error GoTo tratar_erro

txtidcot = 0
txtidcotacao = ""
txtDataemissao = Format(Date, "dd/mm/yy")
txtStatus = "COTANDO"
txtRequisitante = pubUsuario
cmbSetor = pubSetor
Txt_data_validade = "__/__/____"
txtObservacao = ""
CodigoLista = 0
Caption = "Administrativo - Compras - Cotação"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosForn()
On Error GoTo tratar_erro

txtIDListaForn = TBAbrir!ID
txtIDforn = IIf(IsNull(TBAbrir!IDforn), "", TBAbrir!IDforn)
txtforn = IIf(IsNull(TBAbrir!forn), "", TBAbrir!forn)
txtcontatoforn = IIf(IsNull(TBAbrir!contforn), "", TBAbrir!contforn)
txttelforn = IIf(IsNull(TBAbrir!telforn), "", TBAbrir!telforn)
txtfaxforn = IIf(IsNull(TBAbrir!faxforn), "", TBAbrir!faxforn)
txtcondpagtoforn = IIf(IsNull(TBAbrir!CondPagtoforn), "", TBAbrir!CondPagtoforn)
If TBAbrir!CIFforn = True Then Chkcifforn.Value = True Else Chkcifforn.Value = False
If TBAbrir!FOBforn = True Then Chkfobforn.Value = True Else Chkfobforn.Value = False

txtidforn1 = txtIDforn
txtforn1 = txtforn

Frame8.Enabled = True
Novo_Cotacao1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_cot_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lista_cot, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_cot_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_cot.ListItems.Count = 0 Then Exit Sub
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * from compras_cotacao where cotacaotexto = '" & lista_cot.SelectedItem.SubItems(2) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    ProcLimpaCamposCotacao
    ProcCarregaDados
    ProcCarregaListaForn False
    Novo_Cotacao = False
    CodigoLista = lista_cot.SelectedItem.index
End If
TBCotacao.Close
Frame2.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_empenhos
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If ProcVerifSatus("", False) = False Then GoTo Proximo
'                If lista_itens1.SelectedItem.ListSubItems(2) <> "" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_empenhos, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_DblClick()
On Error GoTo tratar_erro

Qtde = 0
qtde_solicitada = ""
With Lista_empenhos
    If .ListItems.Count = 0 Then Exit Sub
    If .SelectedItem.ListSubItems(16) = "FATURADO" Or .SelectedItem.ListSubItems(16) = "FATURADO PARCIAL" Then
        If USMsgBox("Deseja alterar a quantidade empenhada?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            If Alterar = False Then
                USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
Mensagem:
            qtde_solicitada = Txt_quantidade
            qtde_solicitada = InputBox("Favor informar a quantidade empenhada.", , qtde_solicitada)
            If qtde_solicitada = "" Then Exit Sub
            
            If IsNumeric(qtde_solicitada) = False Then
                USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            Qtde = qtde_solicitada
            If Qtde <= 0 Then
                USMsgBox ("So é permitido quantidade maior que 0."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
            
            'Verifica se a quantidade empenhada é maior que a quantidade solicitada
            Qtd = txtQS
            valor = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Qtde_empenho) as Valor from Compras_pedido_lista_empenhos where IDcarteira = " & .SelectedItem.ListSubItems(1) & " and ID <> " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            If Qtd < (valor + Qtde) Then
                USMsgBox ("A quantidade empenhada não pode ser maior que a quantidade cotada, favor alterar."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
                        
            NovoValor = Replace(Qtde, ",", ".")
            Conexao.Execute "Update Compras_pedido_lista_empenhos Set Qtde_empenho = " & NovoValor & " where ID = " & .SelectedItem
            
            USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Empenhar produto/serviço"
            ID_documento = .SelectedItem
            Documento = "Nº cotação: " & txtidcotacao & " - Cód. interno: " & txtdesenho
            Documento1 = "Pedido int.: " & .SelectedItem.ListSubItems(2) & " - Rev.: " & .SelectedItem.ListSubItems(2) & " - Cód. interno: " & Lista.SelectedItem.ListSubItems(4) & " - Rev.: " & .SelectedItem.ListSubItems(3) & " - Qtde. empenhada: " & .SelectedItem.ListSubItems(9) & " - Qtde. entrada: " & .SelectedItem.ListSubItems(10)
            ProcGravaEvento
            '==================================
            ProcCarregaListaEmpenhos
        Else
            ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), .SelectedItem.ListSubItems(5), False
        End If
    Else
        ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), .SelectedItem.ListSubItems(5), False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_empenhos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_empenhos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If ProcVerifSatus("excluir este empenho", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
'            If lista_itens1.SelectedItem.ListSubItems(2) <> "" Then
'                usMsgbox ("Não é permitido excluir este empenho, pois o mesmo está vinculado a uma solicitação."), vbExclamation, "CAPRIND v5.0"
'                .ListItems.Item(InitFor).Checked = False
'            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_forn_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lista_forn
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If ProcVerifSatus("", False) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lista_forn, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_forn_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lista_forn
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If ProcVerifSatus("excluir este fornecedor", True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_forn_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_forn.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from cotacao_fornecedor where id = " & lista_forn.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposForn
    ProcLimpaCamposItem
    ProcCarregaDadosForn
    CodigoLista1 = lista_forn.SelectedItem.index
    Frame4.Enabled = False
End If
TBAbrir.Close

ProcCarregaListaItens
ProcCarregaListaItens1
SSTab3.Tab = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaItens()
On Error GoTo tratar_erro

If txtidcotacao.Text <> "" Then
    Lista_itens.ListItems.Clear
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select CPL.*, CF.* FROM (Cotacao_item CI INNER JOIN compras_pedido_lista CPL ON CI.iditemlista = CPL.idlista) INNER JOIN Cotacao_fornecedor CF ON CF.IDitem = CI.ID where CI.idcot = " & txtidcot & " and CF.IDforn = " & IIf(txtIDforn = "", 0, txtIDforn) & " order by CPL.idlista", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBCompras_Lista.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBCompras_Lista.EOF = False
            ProcCarregaListaItensPadrao Lista_itens
            TBCompras_Lista.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBCompras_Lista.Close
End If
Lista_itens.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaItens1()
On Error GoTo tratar_erro

If txtidcotacao.Text <> "" Then
    Lista_itens1.ListItems.Clear
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select compras_pedido_lista.*, Cotacao_fornecedor.* FROM (Cotacao_item INNER JOIN compras_pedido_lista ON Cotacao_item.iditemlista = compras_pedido_lista.idlista) INNER JOIN Cotacao_fornecedor ON Cotacao_fornecedor.IDitem = Cotacao_item.ID where Cotacao_item.idcot = " & txtidcot & " and Cotacao_fornecedor.IDforn = " & IIf(txtIDforn = "", 0, txtIDforn) & " order by compras_pedido_lista.idlista desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        PBLista.Min = 0
        PBLista.Max = TBCompras_Lista.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBCompras_Lista.EOF = False
            ProcCarregaListaItensPadrao Lista_itens1
            TBCompras_Lista.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBCompras_Lista.Close
End If
Lista_itens1.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaItensPadrao(Lista As ListView)
On Error GoTo tratar_erro

With Lista.ListItems.Add(, , TBCompras_Lista!IDitem)
    .SubItems(1) = IIf(IsNull(TBCompras_Lista!IDlista), "", TBCompras_Lista!IDlista)
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Requisicaotexto from Compras_requisicao where ID_requisicao = " & IIf(IsNull(TBCompras_Lista!ID_Requisicao), 0, TBCompras_Lista!ID_Requisicao), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .SubItems(2) = IIf(IsNull(TBAbrir!Requisicaotexto), "", TBAbrir!Requisicaotexto)
    End If
    .SubItems(3) = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
    .SubItems(4) = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
    .SubItems(5) = IIf(IsNull(TBCompras_Lista!Un), "", TBCompras_Lista!Un)
    .SubItems(6) = IIf(IsNull(TBCompras_Lista!Unidade_com), "", TBCompras_Lista!Unidade_com)
    
    If TBCompras_Lista!Un <> TBCompras_Lista!Unidade_com Then valor = FunConversaoFinalUn(TBCompras_Lista!Un, TBCompras_Lista!Unidade_com, TBCompras_Lista!quant_req, TBCompras_Lista!Desenho, True) Else valor = TBCompras_Lista!quant_req
    .SubItems(7) = FunFormataCasasDecimais(4, valor)
    .SubItems(8) = IIf(IsNull(TBCompras_Lista!quant_req), "", Format(TBCompras_Lista!quant_req, "###,##0.0000"))
    
    .SubItems(9) = IIf(IsNull(TBCompras_Lista!precounit), "", FunFormataCasasDecimais(10, TBCompras_Lista!precounit))
    .SubItems(10) = IIf(IsNull(TBCompras_Lista!Desconto), "", TBCompras_Lista!Desconto)
    .SubItems(11) = IIf(IsNull(TBCompras_Lista!ValorDesconto), "0,00000", FunFormataCasasDecimais(10, TBCompras_Lista!ValorDesconto))
    .SubItems(12) = IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), "0,00000", FunFormataCasasDecimais(10, TBCompras_Lista!preco_unitario_desconto))
    .SubItems(13) = IIf(IsNull(TBCompras_Lista!IPI), "", TBCompras_Lista!IPI)
    .SubItems(14) = IIf(IsNull(TBCompras_Lista!ICMS), "", TBCompras_Lista!ICMS)
    .SubItems(15) = IIf(IsNull(TBCompras_Lista!VlrIPI), "", FunFormataCasasDecimais(2, TBCompras_Lista!VlrIPI))
    .SubItems(16) = IIf(IsNull(TBCompras_Lista!PrecoTotal), "", FunFormataCasasDecimais(2, TBCompras_Lista!PrecoTotal))
    
    .SubItems(17) = IIf(IsNull(TBCompras_Lista!prazoentregaforn), "", Format(TBCompras_Lista!prazoentregaforn, "dd/mm/yy"))
    .SubItems(18) = IIf(IsNull(TBCompras_Lista!detalheitem), "", TBCompras_Lista!detalheitem)
    If TBCompras_Lista!Status_Item = "COTANDO" Or TBCompras_Lista!Status_Item = "AGUARDANDO APROVAÇÃO" Or TBCompras_Lista!Status_Item = "RECEBIDO" Or TBCompras_Lista!Status_Item = "CANCELADO" Then
        Status_Item = TBCompras_Lista!Status_Item
    ElseIf TBCompras_Lista!Status_Item = "N_RECEBIDO" Then
            Status_Item = "COMPRADO"
        Else
            Status_Item = "RECEBIDO PARCIAL"
    End If
    .SubItems(19) = Status_Item
    .SubItems(20) = IIf(IsNull(TBCompras_Lista!Ordem), "", TBCompras_Lista!Ordem)
    .SubItems(21) = IIf(IsNull(TBCompras_Lista!OS), "", TBCompras_Lista!OS)
    Aprovado = ""
    Cor = vbBlack
    If TBCompras_Lista!aprovadoforn = True Then
        Aprovado = "SIM"
        Cor = vbBlue 'Aprovado
    ElseIf TBCompras_Lista!naprovadoforn = True Then
            Aprovado = "NÃO"
            Cor = vbRed 'Aprovar
    End If
    .SubItems(22) = Aprovado
    .SubItems(23) = TBCompras_Lista!ID
    
    'Muda cor da lista
    .ForeColor = Cor
    .ListSubItems(1).ForeColor = Cor
    .ListSubItems(2).ForeColor = Cor
    .ListSubItems(3).ForeColor = Cor
    .ListSubItems(4).ForeColor = Cor
    .ListSubItems(5).ForeColor = Cor
    .ListSubItems(6).ForeColor = Cor
    .ListSubItems(7).ForeColor = Cor
    .ListSubItems(8).ForeColor = Cor
    .ListSubItems(9).ForeColor = Cor
    .ListSubItems(10).ForeColor = Cor
    .ListSubItems(11).ForeColor = Cor
    .ListSubItems(12).ForeColor = Cor
    .ListSubItems(13).ForeColor = Cor
    .ListSubItems(14).ForeColor = Cor
    .ListSubItems(15).ForeColor = Cor
    .ListSubItems(16).ForeColor = Cor
    .ListSubItems(17).ForeColor = Cor
    .ListSubItems(18).ForeColor = Cor
    .ListSubItems(19).ForeColor = Cor
    .ListSubItems(20).ForeColor = Cor
    .ListSubItems(21).ForeColor = Cor
    .ListSubItems(22).ForeColor = Cor
    .ListSubItems(23).ForeColor = Cor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaCamposlibera()
On Error GoTo tratar_erro

If txtStatus = "APROVADA" Or txtStatus = "CANCELADA" Then
    USMsgBox ("Não é permitido liberar esta cotação, pois a cotação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    Liberar = False
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select IDlista from compras_pedido_lista where id_cotacao = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
    USMsgBox ("Não é permitido liberar esta cotação, pois a cotação não tem nenhum produto/serviço."), vbExclamation, "CAPRIND v5.0"
    Liberar = False
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_pedido_lista where id_cotacao = " & txtidcot & " and status_item <> 'Cancelado'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from cotacao_item where idcot = " & Cont & " and iditemlista= " & TBAbrir!IDlista & " order by id", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from cotacao_fornecedor where idcot = " & Cont & " and iditem = " & TBItem!ID, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = True Then
                USMsgBox ("Não é permitido liberar esta cotação, todos os produtos/serviços devem possuir fornecedor."), vbExclamation, "CAPRIND v5.0"
                Liberar = False
                TBAbrir.Close
                Exit Sub
            Else
                Do While TBFornecedor.EOF = False
                    NomeCampo = ""
                    If TBFornecedor!aprovadoforn = True Then
                        If IsNull(TBFornecedor!contforn) = True Or IsNull(TBFornecedor!prazoentregaforn) = True Or IsNull(TBFornecedor!CondPagtoforn) = True Then
                            If IsNull(TBFornecedor!contforn) = True Then NomeCampo = "contato"
                            If IsNull(TBFornecedor!prazoentregaforn) = True Or TBFornecedor!prazoentregaforn = "" Then
                                If NomeCampo <> "" Then NomeCampo = NomeCampo & ", prazo de entrega" Else NomeCampo = "prazo de entrega"
                            End If
                            If IsNull(TBFornecedor!CondPagtoforn) = True Or TBFornecedor!CondPagtoforn = "" Then
                                If NomeCampo <> "" Then NomeCampo = NomeCampo & ", condições de pagamento" Else NomeCampo = "condições de pagamento"
                            End If
                        End If
                        If TBFornecedor!contforn = "" Or IsNull(TBFornecedor!contforn) = True Then
                            If NomeCampo <> "" Then NomeCampo = NomeCampo & ", contato" Else NomeCampo = "contato"
                        End If
                        If TBFornecedor!CIFforn = False And TBFornecedor!FOBforn = False Then
                            If NomeCampo <> "" Then NomeCampo = NomeCampo & ", frete" Else NomeCampo = "frete"
                        End If
                        If TBFornecedor!precounit = "" Or IsNull(TBFornecedor!precounit) = True Or TBFornecedor!precounit = "0" Then
                            If NomeCampo <> "" Then NomeCampo = NomeCampo & ", valor unitário" Else NomeCampo = "valor unitário"
                        End If
                    End If
                    
                    If NomeCampo <> "" Then
                        USMsgBox ("Não é permitido liberar esta cotação, exite(m) algum(ns) dado(s) a ser(em) preenchido(s) para o produto/serviço: " & vbCrLf & TBAbrir!Desenho & " - " & TBAbrir!Descricao & vbCrLf & "Fornecedor: " & TBFornecedor!forn & vbCrLf & "Campos: " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
                        Liberar = False
                        TBAbrir.Close
                        Exit Sub
                    End If
                    TBFornecedor.MoveNext
                Loop
            End If
        End If
        TBAbrir.MoveNext
    Loop
Else
    USMsgBox ("Não é permitido liberar esta cotação, pois todos os produtos/serviços estão cancelados."), vbExclamation, "CAPRIND v5.0"
    Liberar = False
    TBAbrir.Close
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAprovar()
On Error GoTo tratar_erro
Dim Condpagto   As String 'OK
Dim Idaprovado  As Integer 'OK
Dim Totalpedido As Double 'OK

ProcVerificaNaoAprovados

'Verifica se existe fornecedor nessa cotação com certificado ou avalicação vencida
Familiatext = ""
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select CMF.Nome_Razao, CMF.Data_venc, CMF.Fornecedor FROM Cotacao_fornecedor CF INNER JOIN Compras_fornecedores CMF ON CMF.IDcliente = CF.idforn where CF.IDcot = " & Cont & " and CF.aprovadoforn = 'True' group by CMF.Nome_Razao, CMF.Data_venc, CMF.Fornecedor", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    Do While TBFornecedor.EOF = False
        If IsNull(TBFornecedor!Data_venc) = True Or TBFornecedor!Data_venc = "" Or IsNull(TBFornecedor!Fornecedor) = True Or TBFornecedor!Fornecedor = "" Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloquear_fornecedores = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If Familiatext = "" Then Familiatext = "Fornecedor: " & TBFornecedor!Nome_Razao & " - Não homologado" Else Familiatext = Familiatext & " | Fornecedor: " & TBFornecedor!Nome_Razao & " - Não homologado"
            End If
            TBAbrir.Close
        Else
            If TBFornecedor!Data_venc < Date And (TBFornecedor!Fornecedor = "A" Or TBFornecedor!Fornecedor = "C") Then
                If TBFornecedor!Fornecedor = "A" Then NomeCampo = " - Avaliação vencida" Else NomeCampo = " -  Certificado vencido"
                If Familiatext = "" Then Familiatext = "Fornecedor: " & TBFornecedor!Nome_Razao & NomeCampo Else Familiatext = Familiatext & " | Fornecedor: " & TBFornecedor!Nome_Razao & NomeCampo
            End If
        End If
        TBFornecedor.MoveNext
    Loop
    If Familiatext <> "" Then
        USMsgBox ("Não é permitido gerar pedido dessa cotação, pois o(s) fornecedor(es) abaixo não está(ão) homologado(s), ou está(ão) com o certificado vencido ou avaliação vencida: " & vbCrLf & Familiatext), vbExclamation, "CAPRIND v5.0"
        TBFornecedor.Close
        Exit Sub
    End If
End If
TBFornecedor.Close

'Alterar dados da cotação
Conexao.Execute "Update compras_cotacao Set Statuscotacao = 'APROVADA' where id_cotacao = " & txtidcot
txtStatus = "APROVADA"

'Verificar fornecedores aprovados na cotação
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select CI.IDitemlista, CF.* FROM Cotacao_item CI INNER JOIN Cotacao_fornecedor CF ON CI.ID = CF.IDitem where CF.IDcot = " & Cont & " and CF.aprovadoforn = 'True' order by CF.idforn, CF.condpagtoforn, CI.iditemlista", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    Do While TBFornecedor.EOF = False
        'Verificar na tabela compras_pedido_lista itens da cotação q estão liberados sem pedido gerado
        Set TBCompras_Lista = CreateObject("adodb.recordset")
        TBCompras_Lista.Open "Select * from compras_pedido_lista where id_cotacao = " & txtidcot & " and idpedido = 0 and status_item <> 'CANCELADO' and IdLista = " & TBFornecedor!iditemlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras_Lista.EOF = False Then
            If Idaprovado <> TBFornecedor!IDforn Or Condpagto <> TBFornecedor!CondPagtoforn Then
                'Gerar numero do pedido de compra
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from compras_pedido where Year(data) = '" & Year(Date) & "' order by IDPedido desc", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Numero = Left(TBAbrir!Pedido, Len(TBAbrir!Pedido) - 3) + 1
                Else
                    Numero = 1
                End If
                TBAbrir.Close
                Ano = Right(Year(Date), 2)
                
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "Select * from Compras_pedido", Conexao, adOpenKeyset, adLockOptimistic
                TBPedido.AddNew
                TBPedido!Data = Date
                TBPedido!Responsavel = pubUsuario
                
VerifNPedido:
                NumeroPedido = Numero & "/" & Ano
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select * from compras_pedido where Pedido = '" & NumeroPedido & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAliquota.EOF = False Then
                    Numero = Numero + 1
                    GoTo VerifNPedido
                End If
                TBAliquota.Close
                
                TBPedido!Pedido = NumeroPedido
                TBPedido!idcotacao = Cont
                TBPedido!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                TBPedido!IDFornecedor = TBFornecedor!IDforn
                Idaprovado = TBFornecedor!IDforn
                
                'Gravar dados do fornecedor
                Set TBCarteira = CreateObject("adodb.recordset")
                TBCarteira.Open "Select * from compras_fornecedores where idcliente = " & TBFornecedor!IDforn, Conexao, adOpenKeyset, adLockOptimistic
                If TBCarteira.EOF = False Then
                    TBPedido!Fornecedor = TBCarteira!Nome_Razao
                    Fornecedor = TBCarteira!Nome_Razao
                    TBPedido!Categoria = IIf(IsNull(TBCarteira!Categoria), "A", TBCarteira!Categoria)
                    TBPedido!contato = IIf(IsNull(TBFornecedor!contforn), Null, TBFornecedor!contforn)
                    TBPedido!Tipo_endereco = IIf(IsNull(TBCarteira!Tipo_endereco), Null, TBCarteira!Tipo_endereco)
                    TBPedido!Endereco = IIf(IsNull(TBCarteira!Endereco), Null, TBCarteira!Endereco)
                    TBPedido!Numero = IIf(IsNull(TBCarteira!Numero), Null, TBCarteira!Numero)
                    TBPedido!Tipo_bairro = IIf(IsNull(TBCarteira!Tipo_bairro), Null, TBCarteira!Tipo_bairro)
                    TBPedido!Bairro = IIf(IsNull(TBCarteira!Bairro), Null, TBCarteira!Bairro)
                    TBPedido!Cidade = IIf(IsNull(TBCarteira!Cidade), Null, TBCarteira!Cidade)
                    TBPedido!Estado = IIf(IsNull(TBCarteira!Estado), Null, TBCarteira!Estado)
                    TBPedido!Email = IIf(IsNull(TBCarteira!Email), Null, TBCarteira!Email)
                    TBPedido!fone = IIf(IsNull(TBFornecedor!telforn), Null, TBFornecedor!telforn)
                    TBPedido!Fax = IIf(IsNull(TBFornecedor!faxforn), Null, TBFornecedor!faxforn)
                    
                    'Verifica e-mail do contato
                    If IsNull(TBFornecedor!contforn) = False And TBFornecedor!contforn <> "" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from Contatos_fornecedor where IdFornecedor = " & Idaprovado & " and Nome = '" & TBFornecedor!contforn & "' and Email is not null", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBAbrir!Email <> "" Then TBPedido!Email = TBAbrir!Email
                        End If
                        TBAbrir.Close
                    End If
                        
                End If
                TBCarteira.Close
                TBPedido!Status_pedido = "AGUARDANDO APROVAÇÃO"
                USMsgBox ("Cotação aprovada com sucesso. Pedido " & TBPedido!Pedido & " foi gerado para o fornecedor " & Fornecedor & "."), vbInformation, "CAPRIND v5.0"
                TBPedido.Update
                IDpedido = TBPedido!IDpedido
                TBPedido.Close
            End If
            
            TBFornecedor!IDpedido = IDpedido
            TBFornecedor.Update
            
            'Gerar dados comercial
            Set TBAliquota = CreateObject("adodb.recordset")
            TBAliquota.Open "Select * from compras_comercial where IDpedido = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBAliquota.EOF = False Then
                If IsNull(TBFornecedor!Obsforn) = False And TBFornecedor!Obsforn <> "" Then
                    If IsNull(TBAliquota!Observacoes) = False And TBAliquota!Observacoes <> "" Then
                        If TBFornecedor!Obsforn <> TBAliquota!Observacoes Then TBAliquota!Observacoes = TBAliquota!Observacoes & vbCrLf & TBFornecedor!Obsforn
                    Else
                        TBAliquota!Observacoes = TBFornecedor!Obsforn
                    End If
                End If
            Else
                TBAliquota.AddNew
                If IsNull(TBFornecedor!Obsforn) = False And TBFornecedor!Obsforn <> "" Then TBAliquota!Observacoes = TBFornecedor!Obsforn
            End If
            TBAliquota!IDpedido = IDpedido
            TBAliquota!condicoes = TBFornecedor!CondPagtoforn
            Condpagto = TBFornecedor!CondPagtoforn
            TBAliquota.Update
            TBAliquota.Close
            
            'Editar dados do item
            TBCompras_Lista!IDpedido = IDpedido
            TBCompras_Lista!N_referencia = TBFornecedor!N_referencia
            TBCompras_Lista!preco_unitario = Format(TBFornecedor!precounit, "###,##0.0000000000")
            
            ProcAgregarProdutoForn TBCompras_Lista!Codproduto, TBFornecedor!IDforn, TBFornecedor!precounit
            
            'Atualiza valor de compra do produto/serviço
'            ProcAtualizaValorProdServ True, TBFornecedor!Precounit, False, 0, 0, TBCompras_Lista!Desenho
            
            TBCompras_Lista!Desconto = IIf(IsNull(TBFornecedor!Desconto), 0, TBFornecedor!Desconto)
            TBCompras_Lista!ValorDesconto = IIf(IsNull(TBFornecedor!ValorDesconto), 0, Format(TBFornecedor!ValorDesconto, "###,##0.00"))
            TBCompras_Lista!preco_unitario_desconto = IIf(IsNull(TBFornecedor!preco_unitario_desconto), TBFornecedor!precounit, Format(TBFornecedor!preco_unitario_desconto, "###,##0.0000000000"))
            TBCompras_Lista!preco_total = Format(TBFornecedor!PrecoTotal, "###,##0.00")
            TBCompras_Lista!Quant_Comp = TBCompras_Lista!quant_req
            TBCompras_Lista!Quant_Comp_PC = TBCompras_Lista!quant_req_PC
            TBCompras_Lista!Prazo = IIf(IsNull(TBFornecedor!prazoentregaforn), "", TBFornecedor!prazoentregaforn)
            TBCompras_Lista!Status_Item = "AGUARDANDO APROVAÇÃO"
            
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select ID_CF, ID_CFOP from projproduto where Desenho = '" & TBCompras_Lista!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                TBCompras_Lista!ID_CF = TBItem!ID_CF
                TBCompras_Lista!ID_CFOP = TBItem!ID_CFOP
            End If
            TBItem.Close
            
            TBCompras_Lista!IPI = IIf(IsNull(TBFornecedor!IPI), 0, TBFornecedor!IPI)
            TBCompras_Lista!VlrIPI = IIf(IsNull(TBFornecedor!VlrIPI), 0, Format(TBFornecedor!VlrIPI, "###,##0.00"))
            TBCompras_Lista!ICMS = IIf(IsNull(TBFornecedor!ICMS), 0, TBFornecedor!ICMS)
            TBCompras_Lista!vlrICMS = IIf(IsNull(TBFornecedor!vlrICMS), 0, Format(TBFornecedor!vlrICMS, "###,##0.00"))
            TBCompras_Lista!Obs_pedido = IIf(IsNull(TBCompras_Lista!Obs_cotacao), "", TBCompras_Lista!Obs_cotacao)
            
            TBCompras_Lista.Update
            
            ValorTotal = IIf(IsNull(TBCompras_Lista!preco_total), 0, TBCompras_Lista!preco_total)
            NovoValor = Replace(ValorTotal, ",", ".")
            'Calcula valor do centro de custo e adiciona o numero do pedido
            Conexao.Execute "Update compras_pedido_lista_custo Set IDpedido = " & IDpedido & ", valor = (ISNULL(Percentual, 0) * " & NovoValor & ") / 100 where IDlista = " & TBCompras_Lista!IDlista
        End If
        TBCompras_Lista.Close
        TBFornecedor.MoveNext
    Loop
End If
TBFornecedor.Close

ProcGravarTotaisPC IDpedido

'==================================
Modulo = "Compras/Cotação"
Evento = "Gerar pedido"
ID_documento = Cont
Documento = "Nº cotação: " & txtidcotacao
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaListaCotacao (IIf(ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas(2).Caption, Len(lblPaginas(2).Caption) - 5))))
If CodigoLista <> 0 And lista_cot.ListItems.Count <> 0 Then
    lista_cot.SelectedItem = lista_cot.ListItems(CodigoLista)
    lista_cot.SetFocus
End If

Exit Sub
tratar_erro:
    If Err.Number <> "35600" Then USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaCamposGeraPedido()
On Error GoTo tratar_erro

If txtStatus.Text = "APROVADA" Then
    Familiatext = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Pedido from compras_pedido where idcotacao = " & Cont & " and Status_pedido <> 'CANCELADO' order by IDPedido", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If Familiatext <> "" Then Familiatext = Familiatext & "," & TBAbrir!Pedido Else Familiatext = TBAbrir!Pedido
            TBAbrir.MoveNext
        Loop
        USMsgBox ("Pedido(s) gerado(s) para esta cotação: " & vbCrLf & " " & Familiatext), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBAbrir.Close
End If
If USMsgBox("Deseja realmente gerar o(s) pedido(s) de compra(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If txtStatus.Text = "CANCELADA" Then
        USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, pois a cotação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If txtStatus.Text <> "LIBERADA" And txtStatus.Text <> "APROVADA" Then
        USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, pois a cotação ainda não foi liberada."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from compras_pedido_lista where id_cotacao = " & txtidcot & " and idpedido = 0 and Status_Item <> 'CANCELADO' order by idlista", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from Cotacao_fornecedor where IDcot = " & Cont & " and aprovadoforn = 'False' and naprovadoforn = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, todos os produtos/serviços devem ser classificados como aprovados ou não aprovados."), vbExclamation, "CAPRIND v5.0"
                TBFornecedor.Close
                Exit Sub
            End If
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from Cotacao_fornecedor where IDcot = " & Cont & " and aprovadoforn = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = True Then
                USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, pelo menos um dos produtos/serviços deve ser classificado como aprovado."), vbExclamation, "CAPRIND v5.0"
                TBFornecedor.Close
                Exit Sub
            End If
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select CF.* FROM (Cotacao_item CTI INNER JOIN Cotacao_fornecedor CTF ON CTI.ID = CTF.IDitem) INNER JOIN Compras_fornecedores CF ON CF.IDcliente = CTF.IDforn where CTI.iditemlista = " & TBAbrir!IDlista & " and CTI.IDcot = " & Cont & " and CTF.aprovadoforn = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                Do While TBFornecedor.EOF = False
                    If TBFornecedor!Prospecto = True Then
                        USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, pois o fornecedor " & TBFornecedor!Nome_Razao & " está cadastrado como prospecto."), vbExclamation, "CAPRIND v5.0"
                        TBFornecedor.Close
                        Exit Sub
                    End If
                    If TBFornecedor!Endereco = "" Then
                        USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, é necessário informar o endereço do fornecedor " & TBFornecedor!Nome_Razao & " no cadastro de fornecedores."), vbExclamation, "CAPRIND v5.0"
                        TBFornecedor.Close
                        Exit Sub
                    End If
                    If TBFornecedor!Cidade = "" Then
                        USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, é necessário informar a cidade do fornecedor " & TBFornecedor!Nome_Razao & " no cadastro de fornecedores."), vbExclamation, "CAPRIND v5.0"
                        TBFornecedor.Close
                        Exit Sub
                    End If
                    If TBFornecedor!Categoria = "" Then
                        USMsgBox ("Não é permitido gerar o(s) pedido(s) de compra desta cotação, é necessário informar a categoria do fornecedor " & TBFornecedor!Nome_Razao & " no cadastro de fornecedores."), vbExclamation, "CAPRIND v5.0"
                        TBFornecedor.Close
                        Exit Sub
                    End If
                    TBFornecedor.MoveNext
                Loop
            End If
            TBFornecedor.Close
            TBAbrir.MoveNext
        Loop
        ProcAprovar
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaNaoAprovados()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Cotacao_fornecedor.IDitem FROM (Cotacao_item INNER JOIN Cotacao_fornecedor ON Cotacao_item.ID = Cotacao_fornecedor.IDitem) INNER JOIN compras_pedido_lista ON compras_pedido_lista.IDlista = Cotacao_item.iditemlista where Cotacao_fornecedor.IDcot = " & Cont & " and Cotacao_fornecedor.naprovadoforn = 'True' Group by Cotacao_fornecedor.IDitem", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select compras_pedido_lista.* FROM (Cotacao_item INNER JOIN Cotacao_fornecedor ON Cotacao_item.ID = Cotacao_fornecedor.IDitem) INNER JOIN compras_pedido_lista ON compras_pedido_lista.IDlista = Cotacao_item.iditemlista where Cotacao_fornecedor.IDcot = " & Cont & " and Cotacao_fornecedor.IDitem = " & TBAbrir!IDitem & " and Cotacao_fornecedor.aprovadoforn = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = True Then
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select compras_pedido_lista.* FROM (Cotacao_item INNER JOIN Cotacao_fornecedor ON Cotacao_item.ID = Cotacao_fornecedor.IDitem) INNER JOIN compras_pedido_lista ON compras_pedido_lista.IDlista = Cotacao_item.iditemlista where Cotacao_fornecedor.IDcot = " & Cont & " and Cotacao_fornecedor.IDitem = " & TBAbrir!IDitem, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                Do While TBFornecedor.EOF = False
                    If TBFornecedor!ID_Requisicao <> 0 Then
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
                        TBGravar.AddNew
                        TBGravar!ID_Requisicao = TBFornecedor!ID_Requisicao
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select * from Projproduto where Desenho = '" & TBFornecedor!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then
                            TBGravar!Codproduto = TBProduto!Codproduto
                            If TBProduto!Tipo = "P" Or TBProduto!Tipo = "I" Then TBGravar!Tipo = "P" Else TBGravar!Tipo = "S"
                        End If
                        TBProduto.Close
                        TBGravar!CODIGO = TBFornecedor!CODIGO
                        TBGravar!Status_Item = "REQUISIT."
                        TBGravar!Un = TBFornecedor!Un
                        TBGravar!Familia = TBFornecedor!Familia
                        TBGravar!solicitado = TBFornecedor!solicitado
                        TBGravar!setorsolic = TBFornecedor!setorsolic
                        TBGravar!Descricao = TBFornecedor!Descricao
                        TBGravar!quant_req = TBFornecedor!quant_req
                        TBGravar!quant_req_PC = TBFornecedor!quant_req_PC
                        TBGravar!Desenho = TBFornecedor!Desenho
                        TBGravar!N_referencia = TBFornecedor!N_referencia
                        TBGravar!detalheitem = TBFornecedor!detalheitem
                        TBGravar!prazoreq = IIf(IsNull(TBFornecedor!prazoreq), Null, TBFornecedor!prazoreq)
                        TBGravar!Obs = TBFornecedor!Obs
                        TBGravar!Ordem = IIf(IsNull(TBFornecedor!Ordem), 0, TBFornecedor!Ordem)
                        TBGravar!OS = IIf(IsNull(TBFornecedor!OS), Null, TBFornecedor!OS)
                        TBGravar.Update
                        TBGravar.Close
                    End If
                    TBFornecedor!Status_Item = "NÃO APROVADO"
                    TBFornecedor.Update
                    TBFornecedor.MoveNext
                Loop
            End If
            TBFornecedor.Close
        End If
        TBFI.Close
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcCriarNovoNumero() As String
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_Cotacao where Year (dataemissao) = '" & Year(Date) & "' order by id_cotacao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    Numero = Left(TBAbrir!Cotacaotexto, Len(TBAbrir!Cotacaotexto) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: ProcCriarNovoNumero = "COT-0000" & Numero & "/" & Ano
    Case 2: ProcCriarNovoNumero = "COT-000" & Numero & "/" & Ano
    Case 3: ProcCriarNovoNumero = "COT-00" & Numero & "/" & Ano
    Case 4: ProcCriarNovoNumero = "COT-0" & Numero & "/" & Ano
    Case 5: ProcCriarNovoNumero = "COT-" & Numero & "/" & Ano
End Select

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Lista_itens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_itens
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).SubItems(19) <> "COTANDO" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_itens, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_itens1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_itens1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).SubItems(19) <> "COTANDO" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_itens1, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_itens1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_itens1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If .ListItems.Item(InitFor).SubItems(19) <> "COTANDO" Then
                USMsgBox ("Não é permitido excluir este produto/serviço, pois o mesmo está " & .ListItems.Item(InitFor).SubItems(19) & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_itens1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_itens1.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposItem
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select CPL.IDlista, CPL.Desenho, CPL.Familia, CPL.UN, CPL.Unidade_com, CPL.quant_req, CPL.quant_req_PC, CPL.Descricao, CPL.Descricao_comercial, CPL.Obs_cotacao, CF.* FROM (Cotacao_item CI INNER JOIN compras_pedido_lista CPL ON CI.iditemlista = CPL.idlista) INNER JOIN Cotacao_fornecedor CF ON CI.ID = CF.IDitem where CI.id = " & Lista_itens1.SelectedItem & " and CF.IDforn = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    If TBCompras_Lista!Desconto > 0 Then Chk_desc.Value = 1
    
    Txt_ID_tabela_forn = TBCompras_Lista!ID
    txtiditem = TBCompras_Lista!IDitem
    TXTIDLista = IIf(IsNull(TBCompras_Lista!IDlista), "", TBCompras_Lista!IDlista)
    txtdesenho = IIf(IsNull(TBCompras_Lista!Desenho) = False, TBCompras_Lista!Desenho, "")
    
    ProcCarregaComboCodRef cmbReferencia, "P.desenho = '" & txtdesenho & "'", txtIDforn, "F", True, True
    If IsNull(TBCompras_Lista!N_referencia) = False And TBCompras_Lista!N_referencia <> "" Then cmbReferencia = TBCompras_Lista!N_referencia
    
    If IsNull(TBCompras_Lista!Familia) = False And TBCompras_Lista!Familia <> "" Then Cmb_familia = TBCompras_Lista!Familia
    If IsNull(TBCompras_Lista!Un) = False And TBCompras_Lista!Un <> "" Then Cmb_un = TBCompras_Lista!Un
    If IsNull(TBCompras_Lista!Unidade_com) = False And TBCompras_Lista!Unidade_com <> "" Then Cmb_un_com = TBCompras_Lista!Unidade_com
    Txt_quantidade = IIf(IsNull(TBCompras_Lista!quant_req), "", Format(TBCompras_Lista!quant_req, "###,##0.0000"))
    Txt_quantidade_PC = IIf(IsNull(TBCompras_Lista!quant_req_PC), "", TBCompras_Lista!quant_req_PC)
    txtStatusProd = Lista_itens1.SelectedItem.SubItems(19)
    txtDescricao_item = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
    txtDescricao_comercial = IIf(IsNull(TBCompras_Lista!Descricao_comercial), "", TBCompras_Lista!Descricao_comercial)
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select forn from Cotacao_fornecedor where iditem = " & TBCompras_Lista!IDitem & " and aprovadoforn = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_fornecedor_aprovado = IIf(IsNull(TBAbrir!forn), "", TBAbrir!forn)
    End If
    TBAbrir.Close
    
    Txt_vlr_unit_ultima_compra = FunVerifVlrUnitUltCompra(txtdesenho, TXTIDLista)
    
    txtPrazoentregaforn = IIf(IsNull(TBCompras_Lista!prazoentregaforn), "__/__/____", Format(TBCompras_Lista!prazoentregaforn, "dd/mm/yyyy"))
    txtvalorunitforn = IIf(IsNull(TBCompras_Lista!precounit), "", Format(TBCompras_Lista!precounit, "###,##0.0000000000"))
    txtDesconto = IIf(IsNull(TBCompras_Lista!Desconto), "", TBCompras_Lista!Desconto)
    txtIPI = IIf(IsNull(TBCompras_Lista!IPI), "", TBCompras_Lista!IPI)
    TxtvlrIpi = IIf(IsNull(TBCompras_Lista!VlrIPI), "", Format(TBCompras_Lista!VlrIPI, "###,##0.00"))
    txtICMS = IIf(IsNull(TBCompras_Lista!ICMS), "", TBCompras_Lista!ICMS)
    txtvlrICMS = IIf(IsNull(TBCompras_Lista!vlrICMS), "", Format(TBCompras_Lista!vlrICMS, "###,##0.00"))
    txttotalforn = IIf(IsNull(TBCompras_Lista!PrecoTotal), "", Format(TBCompras_Lista!PrecoTotal, "###,##0.00"))
    txtvalordesconto = IIf(IsNull(TBCompras_Lista!ValorDesconto), "", Format(TBCompras_Lista!ValorDesconto, "###,##0.0000000000"))
    txtvalorunitariodesc = IIf(IsNull(TBCompras_Lista!preco_unitario_desconto), "", Format(TBCompras_Lista!preco_unitario_desconto, "###,##0.0000000000"))
    txtObs = IIf(IsNull(TBCompras_Lista!Obs_cotacao), "", TBCompras_Lista!Obs_cotacao)
    Txt_obs_fornecedor = IIf(IsNull(TBCompras_Lista!Obsforn), "", TBCompras_Lista!Obsforn)
    
    If TBCompras_Lista!aprovadoforn = True Then optSim.Value = True
    If TBCompras_Lista!naprovadoforn = True Then optNao.Value = True
    If TBCompras_Lista!aprovadoforn = True Or TBCompras_Lista!naprovadoforn = True Then Frame4.Enabled = False Else Frame4.Enabled = True
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        ProcBloqueiaCamposItem
    Else
        ProcLiberaCamposItem
    End If
    TBProduto.Close
    CodigoLista2 = Lista_itens1.SelectedItem.index
End If
TBCompras_Lista.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloqueiaCamposItem()
On Error GoTo tratar_erro

With Cmb_familia
    .Locked = True
    .TabStop = False
End With
With Cmb_un
    .Locked = True
    .TabStop = False
End With
'With Cmb_un_com
'    .Locked = True
'    .TabStop = False
'End With
With Txt_quantidade
    .Locked = True
    .TabStop = False
End With
With txtDescricao_item
    .Locked = True
    .TabStop = False
End With
With txtDescricao_comercial
    .Locked = True
    .TabStop = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberaCamposItem()
On Error GoTo tratar_erro

With Cmb_familia
    .Locked = False
    .TabStop = True
End With
With Cmb_un
    .Locked = False
    .TabStop = True
End With
'With Cmb_un_com
'    .Locked = False
'    .TabStop = True
'End With
With Txt_quantidade
    .Locked = False
    .TabStop = True
End With
With txtDescricao_item
    .Locked = False
    .TabStop = True
End With
With txtDescricao_comercial
    .Locked = False
    .TabStop = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_itens_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_itens
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If .ListItems.Item(InitFor).SubItems(19) <> "COTANDO" Then
                USMsgBox ("Não é permitido selecionar, pois o produto/serviço está " & .ListItems.Item(InitFor).SubItems(19) & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_solicitados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_solicitados
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_solicitados, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNecessidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaNecessidade
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaNecessidade, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_PCP_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_vendas_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptFim_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_necess_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab > 1 And txtidcotacao = "" Then
    SSTab1.Tab = 1
    Exit Sub
End If

Cmb_empresa.Visible = False
Select Case SSTab1.Tab
    Case 0: If ListaNecessidade.Visible = True Then ListaNecessidade.SetFocus
    Case 1:
        Cmb_empresa.Visible = True
        lista_cot.SetFocus
    Case 2:
        lista_forn.SetFocus
        If ProcVerifProsseguir = False Then Exit Sub
        
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select dataliberada from compras_cotacao where id_cotacao = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then
            If IsNull(TBCompras!dataliberada) = True Or txtStatus.Text = "CANCELADA" Then
                ProcVerifPermissaoUsuario
                If Permitido = False Then
                    USToolBar4.ButtonState(8) = 5
                ElseIf txtStatus.Text <> "CANCELADA" Then
                        USToolBar4.ButtonState(8) = 0
                End If
            Else
                If txtStatus = "APROVADA" Then USToolBar4.ButtonState(8) = 5 Else USToolBar4.ButtonState(8) = 0
            End If
        End If
        TBCompras.Close

        If txtStatus = "APROVADA" Then
            USToolBar4.ButtonState(9) = 5
            USToolBar4.ButtonState(10) = 5
        Else
            USToolBar4.ButtonState(9) = 0
            USToolBar4.ButtonState(10) = 0
        End If
        ProcCarregaListaForn False
    Case 3:
        Lista_itens1.SetFocus
        If ProcVerifProsseguir = False Then Exit Sub
        If txtIDforn = "" Then
            USMsgBox ("Informe o fornecedor na lista antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 2
            Exit Sub
        End If
        txtidforn1 = txtIDforn
        txtforn1 = txtforn
        ProcCarregaListaItens1
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifProsseguir() As Boolean
On Error GoTo tratar_erro

ProcVerifProsseguir = True
If Novo_Cotacao = True Then
    USMsgBox ("Salve a cotação antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 1
    ProcVerifProsseguir = False
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcVerifPermissaoUsuario()
On Error GoTo tratar_erro

Permitido = True
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * from usuarios where usuario = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = 'Compras/Cotação/Liberar cotação'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        Permitido = False
        TBAcessos.Close
        Exit Sub
    End If
    TBAcessos.Close
    If TBUsuarios!Bloqueado = True Then
        Permitido = False
        TBUsuarios.Close
        Exit Sub
    End If
End If
TBUsuarios.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposForn()
On Error GoTo tratar_erro

txtIDListaForn = 0
txtIDforn = ""
txtforn = ""
txtcontatoforn = ""
txttelforn = ""
txtfaxforn = ""
txtcondpagtoforn = ""
Chkcifforn.Value = False
Chkfobforn.Value = False
CodigoLista1 = 0

txtidforn1 = ""
txtforn1 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposItem()
On Error GoTo tratar_erro

Txt_ID_tabela_forn = 0
txtiditem = 0
txtdesenho = ""
cmbReferencia.Clear
Cmb_familia.ListIndex = -1
Cmb_un.ListIndex = -1
Cmb_un_com.ListIndex = -1
Txt_quantidade = ""
Txt_quantidade_PC = ""
txtStatusProd = ""
txtDescricao_item = ""
txtDescricao_comercial = ""
Txt_fornecedor_aprovado = ""
Txt_vlr_unit_ultima_compra = ""
txtPrazoentregaforn.Text = "__/__/____"
txtvalorunitforn.Text = ""
txttotalforn.Text = ""
txtIPI = ""
TxtvlrIpi = ""
txtICMS = ""
txtvlrICMS = ""
txtObs = ""
Txt_obs_fornecedor = ""
optSim.Value = False
optNao.Value = False
CodigoLista2 = 0
Lista_empenhos.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    Select Case SSTab2.Tab
        Case 0: If ListaNecessidade.Visible = True Then ListaNecessidade.SetFocus
        Case 1: Lista_solicitados.SetFocus
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 3 Then
    Select Case SSTab3.Tab
        Case 0: Lista_itens1.SetFocus
        Case 1:
'            Lista_empenhos.SetFocus
            If txtdesenho = "" Then
                SSTab3.Tab = 0
                Exit Sub
            End If
            Txt_qtde_total_cotada = Txt_quantidade

            'Verifica se é requisição de serviço de terceiro
            If Lista_itens1.SelectedItem.ListSubItems(20) <> "" Then
                USMsgBox ("Não é permitido fazer o empenho, pois este produto/serviço já está empenhado para uma ordem de produção."), vbExclamation, "CAPRIND v5.0"
                SSTab3.Tab = 0
                Exit Sub
            End If
            ProcCarregaListaEmpenhos
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quantidade_Change()
On Error GoTo tratar_erro

If Txt_quantidade <> "" Then
    VerifNumero = Txt_quantidade
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_quantidade = ""
        Txt_quantidade.SetFocus
        Exit Sub
    End If
    If Cmb_un <> Cmb_un_com Then
        Txt_quantidade_est = FunFormataCasasDecimais(4, FunConversaoFinalUn(Cmb_un, Cmb_un_com, Txt_quantidade, txtdesenho, True))
    Else
        Txt_quantidade_est = FunFormataCasasDecimais(4, Txt_quantidade)
    End If
    
    If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
        Txt_quantidade_PC = FunCalculaQtdePC(txtdesenho, Txt_quantidade, True, Cmb_un_com)
    Else
        Txt_quantidade_PC = ""
    End If
Else
    Txt_quantidade_est = ""
    Txt_quantidade_PC = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quantidade_LostFocus()
On Error GoTo tratar_erro

Txt_quantidade = Format(Txt_quantidade, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesconto_Change()
On Error GoTo tratar_erro

If Chk_desc.Value = 1 Then
    If txtDesconto.Text <> "" Then
        VerifNumero = txtDesconto.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtDesconto.Text = ""
            txtDesconto.SetFocus
            Exit Sub
        End If
        If txtDesconto > 100 Then
            USMsgBox ("O desconto não pode ser maior que 100."), vbExclamation, "CAPRIND v5.0"
            txtDesconto = ""
            txtDesconto.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaDesconto
    ProcCalculaValores
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto_LostFocus()
On Error GoTo tratar_erro

If txtDesconto = "" Then txtDesconto = 0
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIcms_Change()
On Error GoTo tratar_erro

If txtICMS.Text <> "" Then
    VerifNumero = txtICMS.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtICMS.Text = ""
        txtICMS.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIcms_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtICMS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDforn_Change()
On Error GoTo tratar_erro

If txtIDforn <> "" Then
    VerifNumero = txtIDforn
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDforn = ""
        txtIDforn.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDforn_LostFocus()
On Error GoTo tratar_erro

IDFornecedor = txtIDforn
txtforn.Text = ""
txttelforn.Text = ""
txtfaxforn.Text = ""
Set TBFornecedor = CreateObject("adodb.recordset")
If Novo_Cotacao2 = True Then
    TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & IIf(IDFornecedor = "", 0, IDFornecedor) & " and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & IIf(IDFornecedor = "", 0, IDFornecedor), Conexao, adOpenKeyset, adLockOptimistic
End If
If TBFornecedor.EOF = False Then
    txtforn.Text = IIf(IsNull(TBFornecedor!Nome_Razao), "", Trim(TBFornecedor!Nome_Razao))
    txttelforn.Text = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
    txtfaxforn.Text = IIf(IsNull(TBFornecedor!Fax), "", TBFornecedor!Fax)
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIPI_Change()
On Error GoTo tratar_erro

If txtIPI.Text <> "" Then
    VerifNumero = txtIPI.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIPI.Text = ""
        txtIPI.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtipi_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtIPI

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change(index As Integer)
On Error GoTo tratar_erro

If txtNreg(index) <> "" Then
    VerifNumero = txtNreg(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg(index) = ""
        txtNreg(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change(index As Integer)
On Error GoTo tratar_erro

If txtPagIr(index) <> "" Then
    VerifNumero = txtPagIr(index)
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr(index) = ""
        txtPagIr(index).SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtprazo_sol_Click()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_necess_Change()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (0)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_sol_Change()
On Error GoTo tratar_erro

ProcLimparCamposListaPagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto_Change()
On Error GoTo tratar_erro

If Chk_valor_desc.Value = 1 Then
    If txtvalordesconto.Text <> "" Then
        VerifNumero = txtvalordesconto.Text
        ProcVerificaNumero
        If VerifNumero = False Then
            txtvalordesconto.Text = ""
            txtvalordesconto.SetFocus
            Exit Sub
        End If
        valor = IIf(txtvalorunitforn = "", 0, txtvalorunitforn)
        Valor_Produto = txtvalordesconto
        If Valor_Produto > valor Then
            USMsgBox ("O valor do desconto não pode ser maior que o valor unitário."), vbExclamation, "CAPRIND v5.0"
            txtvalordesconto = ""
            txtvalordesconto.SetFocus
            Exit Sub
        End If
    End If
    ProcCalculaValorDesconto
    ProcCalculaValores
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalordesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalordesconto_LostFocus()
On Error GoTo tratar_erro

If txtvalordesconto = "" Then txtvalordesconto = 0
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValorDesconto()
On Error GoTo tratar_erro

If txtvalorunitforn.Text <> "" And Txt_quantidade <> "" Then
    If IsNumeric(txtvalorunitforn.Text) = True Then
        quantestoque = txtvalorunitforn.Text
        QuantSolicitado = IIf(txtvalordesconto = "", 0, txtvalordesconto)
        If quantestoque <> 0 Then QuantEmpenho = (QuantSolicitado * 100) / quantestoque Else QuantEmpenho = 0
        txtDesconto.Text = QuantEmpenho
        txtvalorunitariodesc.Text = Format(quantestoque - QuantSolicitado, "###,##0.0000000000")
    Else
        Exit Sub
    End If
    ProcCalculaValores
Else
    txtvalordesconto = "0,00000"
    txtvalorunitariodesc = IIf(txtvalorunitforn = "", "0,00000", txtvalorunitforn)
    txttotalforn = "0,00"
    TxtvlrIpi = "0,00"
    txtvlrICMS = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitforn_Change()
On Error GoTo tratar_erro

If txtvalorunitforn.Text <> "" Then
    VerifNumero = txtvalorunitforn.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtvalorunitforn.Text = ""
        txtvalorunitforn.SetFocus
        Exit Sub
    End If
End If
ProcCalculaValor
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaValor()
On Error GoTo tratar_erro

If txtDesconto <> "" And txtDesconto <> "0" Then valor = IIf(txtvalorunitariodesc = "", 0, txtvalorunitariodesc) Else valor = IIf(txtvalorunitforn = "", 0, txtvalorunitforn)
ValorIPI = IIf(txtIPI = "", 0, txtIPI)
ValorICMS = IIf(txtICMS = "", 0, txtICMS)
quantnovo = IIf(Txt_quantidade = "", 0, Txt_quantidade)
TxtvlrIpi = Format(((valor * quantnovo) * ValorIPI) / 100, "###,##0.00")
txtvlrICMS = Format(((valor * quantnovo) * ValorICMS) / 100, "###,##0.00")
txttotalforn = Format(valor * quantnovo, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaCotacao(Pagina As Integer)
On Error GoTo tratar_erro

If Sql_Cotacao_Localizar = "" Then Exit Sub
lblRegistros(2).Caption = "Nº de registros: 0"
lblPaginas(2).Caption = "Página: 0 de: 0"
lista_cot.ListItems.Clear
Set TBLISTA_Compras_Cotacao = CreateObject("adodb.recordset")
TBLISTA_Compras_Cotacao.Open Sql_Cotacao_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Compras_Cotacao.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

lista_cot.ListItems.Clear
TBLISTA_Compras_Cotacao.PageSize = IIf(txtNreg(2) = "", 30, txtNreg(2))
TBLISTA_Compras_Cotacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Compras_Cotacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Compras_Cotacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Compras_Cotacao.PageSize * (Pagina - 1)), 0), TBLISTA_Compras_Cotacao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Compras_Cotacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With lista_cot.ListItems
        .Add , , TBLISTA_Compras_Cotacao!ID_cotacao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Empresa from Empresa where Codigo = " & IIf(IsNull(TBLISTA_Compras_Cotacao!ID_empresa), 0, TBLISTA_Compras_Cotacao!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        End If
        TBAbrir.Close
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Compras_Cotacao!Cotacaotexto), "", TBLISTA_Compras_Cotacao!Cotacaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Compras_Cotacao!DataEmissao), "", Format(TBLISTA_Compras_Cotacao!DataEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Compras_Cotacao!statuscotacao), "", TBLISTA_Compras_Cotacao!statuscotacao)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Compras_Cotacao!Requisitante), "", TBLISTA_Compras_Cotacao!Requisitante)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Compras_Cotacao!Setor), "", TBLISTA_Compras_Cotacao!Setor)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Compras_Cotacao!ID_empresa), 0, TBLISTA_Compras_Cotacao!ID_empresa)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Compras_Cotacao!Data_validade), "", Format(TBLISTA_Compras_Cotacao!Data_validade, "dd/mm/yy"))
    End With
    TBLISTA_Compras_Cotacao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros(2).Caption = "Nº de registros: " & TBLISTA_Compras_Cotacao.RecordCount
If TBLISTA_Compras_Cotacao.AbsolutePage = adPosBOF Then
   lblPaginas(2).Caption = "Página: 1 de: " & TBLISTA_Compras_Cotacao.PageCount
ElseIf TBLISTA_Compras_Cotacao.AbsolutePage = adPosEOF Then
        lblPaginas(2).Caption = "Página: " & TBLISTA_Compras_Cotacao.PageCount & " de: " & TBLISTA_Compras_Cotacao.PageCount
    Else
        lblPaginas(2).Caption = "Página: " & TBLISTA_Compras_Cotacao.AbsolutePage - 1 & " de: " & TBLISTA_Compras_Cotacao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBCotacao!ID_empresa) = False And TBCotacao!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBCotacao!ID_empresa
Caption = "Compras - Cotação - (Cotação : " & IIf(IsNull(TBCotacao!Cotacaotexto), "", TBCotacao!Cotacaotexto) & ")"
txtidcot = TBCotacao!ID_cotacao
Cont = txtidcot
txtidcotacao = IIf(IsNull(TBCotacao!Cotacaotexto), "", TBCotacao!Cotacaotexto)
txtDataemissao.Text = IIf(IsNull(TBCotacao!DataEmissao), "", Format(TBCotacao!DataEmissao, "dd/mm/yy"))
txtStatus.Text = IIf(IsNull(TBCotacao!statuscotacao), "", TBCotacao!statuscotacao)
txtRequisitante.Text = IIf(IsNull(TBCotacao!Requisitante), "", TBCotacao!Requisitante)
cmbSetor.Text = IIf(IsNull(TBCotacao!Setor), "", TBCotacao!Setor)
Txt_data_validade = IIf(IsNull(TBCotacao!Data_validade), "__/__/____", Format(TBCotacao!Data_validade, "dd/mm/yyyy"))
txtObservacao = IIf(IsNull(TBCotacao!Obs), "", TBCotacao!Obs)
Novo_Cotacao = False
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitforn_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtvalorunitforn

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtvalorunitforn_LostFocus()
On Error GoTo tratar_erro

txtvalorunitforn = Format(txtvalorunitforn, "###,##0.0000000000")
txtvalordesconto = Format(txtvalordesconto, "###,##0.0000000000")
txtvalorunitariodesc = Format(txtvalorunitariodesc, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaDesconto()
On Error GoTo tratar_erro

If txtvalorunitforn.Text <> "" And Txt_quantidade <> "" And txtDesconto <> "" Then
    If IsNumeric(txtvalorunitforn.Text) = True Then
        a = Format(txtvalorunitforn.Text, "###,##0.0000000000")
        c = IIf(txtDesconto = "", 0, txtDesconto)
        D = (a * c) / 100
        txtvalordesconto.Text = Format(D, "###,##0.0000000000")
        txtvalorunitariodesc.Text = Format(a - D, "###,##0.0000000000")
        ProcCalculaValores
    End If
Else
    txtvalordesconto = "0,00000"
    txtvalorunitforn = IIf(txtvalorunitforn = "", "0,00000", txtvalorunitforn)
    txttotalforn = "0,00"
    TxtvlrIpi = "0,00"
    txtvlrICMS = "0,00"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaValores()
On Error GoTo tratar_erro

If Txt_quantidade = "" Or txtvalorunitforn = "" Then Exit Sub
'Atribui valores
If txtvalorunitariodesc.Text = "" Or txtvalorunitariodesc.Text = "0,00000" Then
    txttotalforn = Format(txtvalorunitforn * Txt_quantidade, "###,##0.00")
Else
    txttotalforn = Format(txtvalorunitariodesc * Txt_quantidade, "###,##0.00")
End If
ProcCalculaValor
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarCot()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With ListaNecessidade
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente gerar cotação deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                ProcLimpaCamposCotacao
                ProcLimparTudo
                ProcNovaCotacao
                
                ProcConfVariaveisLocForn False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
                Sit_REG = 1
                FrmCompras_localizafornecedor.Show 1
                
                'Se não adicionou nenhum fornecedor, exclui a cotação e para o código
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select id from Cotacao_fornecedor where idcot = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = True Then
                    Conexao.Execute "DELETE from compras_cotacao where id_cotacao = " & txtidcot
                    ProcLimpaCamposCotacao
                    TBFornecedor.Close
                    Exit Sub
                End If
                TBFornecedor.Close
                
                If USMsgBox("Algum produto/serviço selecionado será adicionado com quantidade parcial?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido1 = True
            End If
            Permitido = True
            
            IDlista = .ListItems.Item(InitFor)
            Desenho = .ListItems(InitFor).SubItems(1)
            If Permitido1 = True Then
                Compras_Pedido = False
                Vendas_PI = False
                Compras_Cotacao = True
                Faturamento = False
                Qtde = .ListItems(InitFor).SubItems(4)
                Permitido2 = True
                Sit_REG = 1
                frmVendas_PI_liberaritem.Show 1
                If Permitido2 = False Then
                    valor = .ListItems(InitFor).SubItems(4)
                    ProcNovo_Necessidade Opt_vendas
                End If
            Else
                valor = .ListItems(InitFor).SubItems(4)
                ProcNovo_Necessidade Opt_vendas
            End If
        End If
    Next InitFor
End With

With Lista_solicitados
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente gerar cotação deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                ProcLimpaCamposCotacao
                ProcLimparTudo
                ProcNovaCotacao
                
                ProcConfVariaveisLocForn False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
                Sit_REG = 1
                FrmCompras_localizafornecedor.Show 1
                
                'Se não adicionou nenhum fornecedor, exclui a cotação e para o código
                Set TBFornecedor = CreateObject("adodb.recordset")
                TBFornecedor.Open "Select id from Cotacao_fornecedor where idcot = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
                If TBFornecedor.EOF = True Then
                    Conexao.Execute "DELETE from compras_cotacao where id_cotacao = " & txtidcot
                    ProcLimpaCamposCotacao
                    TBFornecedor.Close
                    Exit Sub
                End If
                TBFornecedor.Close
            End If
            Permitido = True
            ProcNovo_Solicitacao .ListItems.Item(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de gerar a cotação."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    USMsgBox ("Nova cotação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Frame2.Enabled = True
    Novo_Cotacao = False
    
    ProcCarregalista_Necessidade
    ProcCarregalista_Solicitacao
    
    Sql_Cotacao_Localizar = "Select * from Compras_Cotacao where id_cotacao = " & txtidcot
    ProcCarregaListaCotacao (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovaCotacao()
On Error GoTo tratar_erro

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * from compras_cotacao", Conexao, adOpenKeyset, adLockOptimistic
TBCotacao.AddNew
TBCotacao!statuscotacao = "COTANDO"
TBCotacao!Cotacaotexto = ProcCriarNovoNumero
TBCotacao!ID_empresa = Cmb_empresa_carteira.ItemData(Cmb_empresa_carteira.ListIndex)
TBCotacao!DataEmissao = Date
TBCotacao!Requisitante = pubUsuario
TBCotacao!Setor = pubSetor
'TBCotacao!Data_validade = IIf(Txt_data_validade = "__/__/____", Null, Txt_data_validade)
'TBCotacao!Obs = txtobservacao
TBCotacao.Update
'==================================
Modulo = "Compras/Cotação"
Evento = "Novo"
ID_documento = TBCotacao!ID_cotacao
Documento = "Nº cotação: " & TBCotacao!Cotacaotexto
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaDados
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovo_Necessidade(Necess_PI As Boolean)
On Error GoTo tratar_erro

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM Cotacao_item where coditem = '" & Desenho & "' and Cotacao_item.idcot = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from compras_pedido_lista where IDlista = " & TBCotacao!iditemlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
Else
    TBCotacao.AddNew
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
    TBGravar.AddNew
End If

TBGravar!ID_cotacao = txtidcot
TBGravar!IDpedido = 0

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select * from projproduto where codproduto = " & IDlista & "", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBGravar!Codproduto = IIf(IsNull(TBProduto!Codproduto), "", TBProduto!Codproduto)
    TBGravar!Desenho = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
    TBGravar!Descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    TBGravar!Descricao_comercial = IIf(IsNull(TBProduto!descricaotecnica), "", TBProduto!descricaotecnica)
    TBGravar!Un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
    TBGravar!Unidade_com = IIf(IsNull(TBProduto!Unidade_com), "", TBProduto!Unidade_com)
    TBGravar!Familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
    If TBProduto!Tipo = "S" Then TBGravar!Tipo = "S" Else TBGravar!Tipo = "P"
    
    If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
        TBGravar!Quant_Comp_PC = FunCalculaQtdePC(TBProduto!Desenho, valor, True, TBProduto!Unidade_com)
        TBGravar!quant_req_PC = FunCalculaQtdePC(TBProduto!Desenho, valor, True, TBProduto!Unidade_com)
    End If
End If
TBProduto.Close

TBGravar!Quant_Comp = 0
TBGravar!quant_req = valor

'Calcula quantidade se a unidade for diferente
If TBGravar!Un <> TBGravar!Unidade_com Then
    If FunVerifUNConversao(TBGravar!Un, TBGravar!Unidade_com) = True Then
        TBGravar!Qtde_estoque = FunConverteUN(TBGravar!Un, TBGravar!Unidade_com, valor, TBGravar!Desenho)
    Else
        TBGravar!Qtde_estoque = valor / FunVerificaTabelaConversaoUnidade(TBGravar!Un, TBGravar!Unidade_com)
    End If
Else
    TBGravar!Qtde_estoque = Null
End If

TBGravar!Status_Item = "COTANDO"
TBGravar!ValorDesconto = 0
TBGravar.Update

TBCotacao!idcot = txtidcot
TBCotacao!iditemlista = TBGravar!IDlista
TBCotacao!coditem = Desenho
TBCotacao.Update

If Necess_PI = True Then
    Valor3 = valor
    'Empenha a cotação para os pedidos de venda mais antigos
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select ENDT.ID, ENDT.Requisitado - ISNULL(CPLE.Qtde_empenho, 0) AS Requisitado from Estoque_necessidade_detalhado ENDT LEFT JOIN Compras_pedido_lista_empenhos CPLE ON CPLE.IDcarteira = ENDT.ID where ENDT.Desenho = '" & Desenho & "' and ENDT.Tipo <> 'OP' and ENDT.Requisitado > ISNULL(CPLE.Qtde_empenho, 0) order by ENDT.Prazo", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False And Valor3 > 0
            Set TBCotacao = CreateObject("adodb.recordset")
            TBCotacao.Open "Select * FROM Compras_pedido_lista_empenhos", Conexao, adOpenKeyset, adLockOptimistic
            TBCotacao.AddNew
            TBCotacao!Data = Date
            TBCotacao!Responsavel = pubUsuario
            TBCotacao!IDlista = TBGravar!IDlista
            TBCotacao!IDcarteira = TBCFOP!ID
            If Valor3 >= TBCFOP!Requisitado Then
                TBCotacao!Qtde_empenho = TBCFOP!Requisitado
                Valor3 = Valor3 - TBCFOP!Requisitado
            Else
                TBCotacao!Qtde_empenho = Valor3
                Valor3 = 0
            End If
            TBCotacao.Update
            TBCFOP.MoveNext
        Loop
    End If
End If

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM Cotacao_item where coditem = '" & Desenho & "' and Idcot = " & txtidcot, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    Evento = "Novo produto/serviço"
    ProcGravaFornecedores TBCotacao!ID, TBGravar!IDlista, Desenho
End If
TBCotacao.Close
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovo_Solicitacao(IDlista As Long)
On Error GoTo tratar_erro

Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from Compras_pedido_lista where IDlista = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Cont = txtidcot
    TBFamilia!Status_Item = "COTANDO"
    TBFamilia!ID_cotacao = Cont
    TBFamilia!IDpedido = 0
    TBFamilia!Obs_cotacao = IIf(IsNull(TBFamilia!Obs), "", TBFamilia!Obs)
    TBFamilia.Update
    
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from cotacao_item where iditemlista = " & TBFamilia!IDlista & " and idcot = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Evento = "Alterar produto/serviço"
    Else
        TBItem.AddNew
        Evento = "Novo produto/serviço"
    End If
    
    TBItem!idcot = Cont
    TBItem!iditemlista = TBFamilia!IDlista
    TBItem!coditem = TBFamilia!Desenho
    TBItem.Update
    ProcGravaFornecedores TBItem!ID, TBFamilia!IDlista, TBFamilia!Desenho
    TBItem.Close
End If
TBFamilia.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravaFornecedores(IDitem As Long, IDlista As Long, Codinterno As String)
On Error GoTo tratar_erro

With frmcompras_reqcot.lista_forn
    For InitFor1 = 1 To .ListItems.Count
        If .ListItems.Item(InitFor1).Checked = True Then
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select * from Cotacao_fornecedor where idcot = " & Cont & " and idforn = " & .ListItems.Item(InitFor1).ListSubItems(8).Text, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                If TBFornecedor!IDitem = 0 Then
                    TBFornecedor!IDitem = IDitem
                    
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select IA.n_referencia from item_aplicacoes IA INNER JOIN Projproduto P ON P.Codproduto = IA.Codproduto where P.Desenho = '" & Codinterno & "' and IA.ID_cliente_forn = " & TBFornecedor!IDforn & " and IA.Tipo = 'F'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        TBFornecedor!N_referencia = TBAfericao!N_referencia
                    End If
                    TBAfericao.Close
                    
                    TBFornecedor.Update
                Else
                    Set TBTempo = CreateObject("adodb.recordset")
                    TBTempo.Open "Select * from Cotacao_fornecedor where idcot = " & Cont & " and idforn = " & TBFornecedor!IDforn & " and iditem = " & IDitem, Conexao, adOpenKeyset, adLockOptimistic
                    If TBTempo.EOF = True Then TBTempo.AddNew
                    TBTempo!IDforn = TBFornecedor!IDforn
                    TBTempo!forn = TBFornecedor!forn
                    TBTempo!idcot = TBFornecedor!idcot
                    TBTempo!contforn = TBFornecedor!contforn
                    TBTempo!telforn = TBFornecedor!telforn
                    TBTempo!faxforn = TBFornecedor!faxforn
                    TBTempo!CondPagtoforn = TBFornecedor!CondPagtoforn
                    TBTempo!CIFforn = TBFornecedor!CIFforn
                    TBTempo!FOBforn = TBFornecedor!FOBforn
                    TBTempo!aprovadoforn = False
                    TBTempo!naprovadoforn = False
                    TBTempo!IDitem = IDitem
                    
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select IA.n_referencia from item_aplicacoes IA INNER JOIN Projproduto P ON P.Codproduto = IA.Codproduto where P.Desenho = '" & Codinterno & "' and IA.ID_cliente_forn = " & TBFornecedor!IDforn & " and IA.Tipo = 'F'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        TBTempo!N_referencia = TBAfericao!N_referencia
                    End If
                    TBAfericao.Close
                    
                    TBTempo.Update
                    TBTempo.Close
                End If
            End If
            TBFornecedor.Close
            
            '==================================
            Modulo = "Compras/Cotação"
            ID_documento = IDlista
            Documento = "Nº cotação: " & txtidcotacao
            Documento1 = "Fornecedor: " & .ListItems.Item(InitFor1).ListSubItems(1).Text & " - Cód. interno: " & Codinterno
            ProcGravaEvento
            '==================================
            
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select Codproduto from Projproduto where Desenho = '" & Codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                ProcAgregarProdutoForn TBProduto!Codproduto, .ListItems.Item(InitFor1).ListSubItems(8).Text, 0
            End If
            TBProduto.Close
            
        End If
    Next InitFor1
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoEmpenho()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select IDlista from Compras_pedido_lista where IDlista = " & TXTIDLista & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Sit_REG = 0 Else Sit_REG = 1
TBAbrir.Close

Compras_Requisicao = False
Compras_Cotacao = True
Compras_Pedido = False
frmProd_Lista_Produto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpenho()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Permitido = False
With Lista_empenhos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where ID = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Excluir empenho do produto/serviço"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº cotação: " & txtidcotacao & " - Cód. interno: " & txtdesenho
            Documento1 = "Pedido int.: " & .ListItems(InitFor).ListSubItems(2) & " - Rev.: " & .ListItems(InitFor).ListSubItems(3) & " - Cód. interno: " & .ListItems(InitFor).ListSubItems(5) & " - Rev.: " & .ListItems(InitFor).ListSubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaEmpenhos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaEmpenhos()
On Error GoTo tratar_erro

Valor3 = 0
Lista_empenhos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VC.*, CPLE.ID, CPLE.Qtde_empenho, CPLE.Qtde_recebida FROM vendas_carteira VC INNER JOIN Compras_pedido_lista_empenhos CPLE on VC.codigo = CPLE.IDCarteira where CPLE.IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_empenhos.ListItems.Add(, , TBLISTA!ID)
            .SubItems(1) = TBLISTA!CODIGO
            
            Set TBCFOP = CreateObject("adodb.recordset")
            If IsNull(TBLISTA!ID_solicitacao) = True Or TBLISTA!ID_solicitacao = 0 Then
                TBCFOP.Open "Select Ncotacao, Revisao, cliente FROM vendas_proposta where cotacao = " & TBLISTA!Cotacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    .SubItems(2) = IIf(IsNull(TBCFOP!Ncotacao), "", TBCFOP!Ncotacao)
                    .SubItems(3) = IIf(IsNull(TBCFOP!Revisao), "", TBCFOP!Revisao)
                    .SubItems(4) = IIf(IsNull(TBCFOP!Cliente), "", TBCFOP!Cliente)
                End If
            Else
                TBCFOP.Open "Select Requisicaotexto FROM Outros_SolicitacaoPCP where ID = " & TBLISTA!ID_solicitacao, Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then .SubItems(2) = IIf(IsNull(TBCFOP!Requisicaotexto), "", TBCFOP!Requisicaotexto)
            End If
            TBCFOP.Close
            
            .SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .SubItems(6) = IIf(IsNull(TBLISTA!Rev_codinterno), "", TBLISTA!Rev_codinterno)
            .SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .SubItems(8) = IIf(IsNull(TBLISTA!descricao_tecnica), "", TBLISTA!descricao_tecnica)
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .SubItems(9) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_recebida), 0, TBLISTA!Qtde_recebida)
            .SubItems(10) = Format(Valor1, "###,##0.0000")
            .SubItems(11) = Format(valor - Valor1, "###,##0.0000")
            .SubItems(12) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .SubItems(13) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
            .SubItems(14) = IIf(IsNull(TBLISTA!PCCliente), "", TBLISTA!PCCliente)
            .SubItems(15) = IIf(IsNull(TBLISTA!N_item), "", TBLISTA!N_item)
            .SubItems(16) = IIf(IsNull(TBLISTA!Liberacao), "", TBLISTA!Liberacao)
        End With
        Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close
valor = Txt_qtde_total_cotada
Txt_qtde_total_emp = Format(Valor3, "###,##0.0000")
Txt_qtde_total_disp = Format(valor - Valor3, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1:
        If SSTab2.Tab = 0 Then ProcFiltrar_Necessidade Else ProcFiltrar_Solicitacao
    Case 2: ProcGerarCot
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcLiberar
    Case 9: procEmitir_PI
    Case 10: procAtualiza
    Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procNovoForn
    Case 2: procSalvarForn
    Case 3: procExcluirForn
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: procAdd_item
    Case 8: procAprovarForn
    Case 9: procNao_aprovar
    Case 10: procCancelar_Aprovacao
    Case 12: ProcAjuda
    Case 13: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procSalvar_item
    Case 2: procExcluir_item
    Case 3: ProcImprimir
    Case 4: procAnteriorForn
    Case 5: procProximoForn
    Case 6: procExcluirTodosForn
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar6_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoEmpenho
    Case 2: ProcExcluirEmpenho
    Case 3: ProcImprimir
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifSatus(Acao As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifSatus = True
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select statuscotacao from Compras_Cotacao where id_cotacao = " & txtidcot & " and statuscotacao <> 'COTANDO'", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & ", pois a cotação está " & IIf(IsNull(TBFI!statuscotacao), "", TBFI!statuscotacao) & "."), vbExclamation, "CAPRIND v5.0"
    ProcVerifSatus = False
End If
TBFI.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcLimparCamposListaPagina(IndexTexto As Integer)
On Error GoTo tratar_erro

If IndexTexto = 0 Then ListaNecessidade.ListItems.Clear Else Lista_solicitados.ListItems.Clear
lblRegistros(IndexTexto).Caption = "Nº de registros: 0"
lblPaginas(IndexTexto).Caption = "Página: 0 de: 0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
