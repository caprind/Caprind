VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Requisicao 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outros - Solicitação"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   103
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
      SearchText      =   "Atualizando..."
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
      ItemData        =   "frmCompras_Requisicao.frx":0000
      Left            =   270
      List            =   "frmCompras_Requisicao.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1705
      Width           =   4860
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17754
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
      TabCaption(0)   =   "Solicitação de compra"
      TabPicture(0)   =   "frmCompras_Requisicao.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lista_req"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Cmd_dados_cancelamento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Txt_ID_req"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ActiveResize1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Lista de produtos/serviços"
      TabPicture(1)   =   "frmCompras_Requisicao.frx":0020
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   -75000
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
      Begin VB.TextBox Txt_ID_req 
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
         Left            =   -72780
         TabIndex        =   112
         Text            =   "0"
         ToolTipText     =   "IDLista."
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   106
         Top             =   9090
         Width           =   15195
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
            ItemData        =   "frmCompras_Requisicao.frx":003C
            Left            =   7020
            List            =   "frmCompras_Requisicao.frx":004C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   180
            Width           =   1965
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
            TabIndex        =   15
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
            TabIndex        =   13
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   19
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Requisicao.frx":0077
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
            TabIndex        =   18
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Requisicao.frx":381B
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
            TabIndex        =   16
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
         Begin DrawSuite2022.USButton cmdPagUlt 
            Height          =   315
            Left            =   12300
            TabIndex        =   20
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Requisicao.frx":7324
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
         Begin DrawSuite2022.USButton cmdPagPrim 
            Height          =   315
            Left            =   10680
            TabIndex        =   17
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_Requisicao.frx":ABB0
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
            Left            =   3510
            TabIndex        =   125
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
            Left            =   5670
            TabIndex        =   114
            Top             =   240
            Width           =   1260
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
            TabIndex        =   109
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
            TabIndex        =   108
            Top             =   240
            Width           =   1275
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
            TabIndex        =   107
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.CommandButton Cmd_dados_cancelamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71220
         Picture         =   "frmCompras_Requisicao.frx":EC9F
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Verificar dados do cancelamento."
         Top             =   2310
         Width           =   315
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   -74970
         TabIndex        =   60
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
            MouseIcon       =   "frmCompras_Requisicao.frx":EDA1
            MousePointer    =   99  'Custom
            TabIndex        =   65
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
            MouseIcon       =   "frmCompras_Requisicao.frx":F0AB
            MousePointer    =   99  'Custom
            TabIndex        =   64
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
            MouseIcon       =   "frmCompras_Requisicao.frx":F3B5
            MousePointer    =   99  'Custom
            TabIndex        =   63
            ToolTipText     =   "Departamento do contato."
            Top             =   630
            Width           =   9855
         End
         Begin VB.TextBox txttelcontato 
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
            MaxLength       =   40
            MouseIcon       =   "frmCompras_Requisicao.frx":F6BF
            MousePointer    =   99  'Custom
            TabIndex        =   62
            ToolTipText     =   "Ramal do contato."
            Top             =   1020
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
            MouseIcon       =   "frmCompras_Requisicao.frx":F9C9
            MousePointer    =   99  'Custom
            TabIndex        =   61
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
            Top             =   1478
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2595
         Left            =   -74925
         TabIndex        =   70
         Top             =   1320
         Width           =   15195
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
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   990
            Width           =   3675
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
            Left            =   4110
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   990
            Width           =   1755
         End
         Begin VB.TextBox txtData_Solicitacao 
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
            Left            =   6510
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1005
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
            Left            =   9570
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da aprovação."
            Top             =   990
            Width           =   1755
         End
         Begin VB.TextBox txtSolicitado 
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
            Left            =   7530
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   3600
         End
         Begin VB.TextBox txtnumero 
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
            Left            =   5070
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Número da solicitação de compra."
            Top             =   390
            Width           =   1425
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
            Left            =   11340
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela aprovação."
            Top             =   990
            Width           =   3675
         End
         Begin VB.TextBox txtSetor_Solicitado 
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
            Left            =   11150
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Setor."
            Top             =   390
            Width           =   3865
         End
         Begin VB.TextBox txtobssolicitacao 
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
            TabIndex        =   11
            ToolTipText     =   "Observações da solicitação."
            Top             =   1620
            Width           =   14835
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
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Status da solicitação."
            Top             =   990
            Width           =   3495
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável pela aprovação"
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
            Left            =   12142
            TabIndex        =   119
            Top             =   780
            Width           =   2070
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data/hora aprovação"
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
            Left            =   9675
            TabIndex        =   118
            Top             =   780
            Width           =   1545
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
            Left            =   2040
            TabIndex        =   80
            Top             =   180
            Width           =   735
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
            Left            =   8873
            TabIndex        =   79
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
            Left            =   6840
            TabIndex        =   78
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label9 
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
            Left            =   6727
            TabIndex        =   77
            Top             =   780
            Width           =   1980
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
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
            Left            =   4260
            TabIndex        =   76
            Top             =   780
            Width           =   1455
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
            TabIndex        =   75
            Top             =   4200
            Width           =   270
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nº solicitação"
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
            Left            =   5212
            TabIndex        =   74
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label14 
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
            Left            =   12887
            TabIndex        =   73
            Top             =   180
            Width           =   390
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
            Left            =   7162
            TabIndex        =   72
            Top             =   1410
            Width           =   870
         End
         Begin VB.Label Label23 
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
            Left            =   1650
            TabIndex        =   71
            Top             =   780
            Width           =   555
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   102
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   16
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   353
         ButtonTop8      =   2
         ButtonWidth8    =   45
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Copiar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Copiar (F8)"
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
         ButtonLeft9     =   400
         ButtonTop9      =   2
         ButtonWidth9    =   44
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Validação"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Validar/Cancelar validação (F9)"
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
         ButtonLeft10    =   446
         ButtonTop10     =   2
         ButtonWidth10   =   62
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Aprovação"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Aprovar/cancelar aprovação (F10)"
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
         ButtonLeft11    =   510
         ButtonTop11     =   2
         ButtonWidth11   =   69
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Atualizar"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft12    =   581
         ButtonTop12     =   2
         ButtonWidth12   =   59
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonAlignment13=   2
         ButtonType13    =   1
         ButtonStyle13   =   -1
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   -1
         ButtonLeft13    =   642
         ButtonTop13     =   4
         ButtonWidth13   =   2
         ButtonHeight13  =   54
         ButtonCaption14 =   "Ajuda"
         ButtonEnabled14 =   0   'False
         ButtonIconSize14=   32
         ButtonToolTipText14=   "Ajuda (F1)"
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
         ButtonLeft14    =   646
         ButtonTop14     =   2
         ButtonWidth14   =   41
         ButtonHeight14  =   21
         ButtonUseMaskColor14=   0   'False
         ButtonCaption15 =   "Sair"
         ButtonEnabled15 =   0   'False
         ButtonIconSize15=   32
         ButtonToolTipText15=   "Sair (Esc)"
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
         ButtonLeft15    =   689
         ButtonTop15     =   2
         ButtonWidth15   =   30
         ButtonHeight15  =   21
         ButtonUseMaskColor15=   0   'False
         ButtonEnabled16 =   0   'False
         ButtonIconSize16=   32
         ButtonKey16     =   "16"
         ButtonAlignment16=   2
         BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState16   =   5
         ButtonLeft16    =   721
         ButtonTop16     =   2
         ButtonWidth16   =   24
         ButtonHeight16  =   24
         ButtonUseMaskColor16=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12090
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Requisicao.frx":FCD3
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   75
         TabIndex        =   104
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
         ButtonCaption7  =   "Status"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Status (F7)"
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
         ButtonLeft7     =   309
         ButtonTop7      =   2
         ButtonWidth7    =   45
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Centro de custo"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar centro de custo"
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
         ButtonLeft8     =   356
         ButtonTop8      =   2
         ButtonWidth8    =   85
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
         ButtonLeft9     =   443
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   447
         ButtonTop10     =   2
         ButtonWidth10   =   41
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
         ButtonLeft11    =   490
         ButtonTop11     =   2
         ButtonWidth11   =   30
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
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
         ButtonState12   =   5
         ButtonLeft12    =   522
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   8970
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_Requisicao.frx":19427
            Count           =   1
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8835
         Left            =   75
         TabIndex        =   81
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   15584
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
         TabCaption(0)   =   "Dados dos produtos/serviços"
         TabPicture(0)   =   "frmCompras_Requisicao.frx":1FFAC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "PBLista1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lista"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Framelista"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtidcarteira"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtIDLista"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame4"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Txt_ID_PC"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtQS_com_PC"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Centro de custo"
         TabPicture(1)   =   "frmCompras_Requisicao.frx":1FFC8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame8"
         Tab(1).Control(1)=   "txtIDCentro"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Lista_custo"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Empenhos"
         TabPicture(2)   =   "frmCompras_Requisicao.frx":1FFE4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Lista_empenhos"
         Tab(2).Control(1)=   "Frame6"
         Tab(2).ControlCount=   2
         Begin VB.TextBox txtQS_com_PC 
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
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   126
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade solicitada em peça."
            Top             =   6540
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   -74940
            TabIndex        =   120
            Top             =   7530
            Width           =   15075
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
               TabIndex        =   58
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade disponível."
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
               TabIndex        =   57
               TabStop         =   0   'False
               ToolTipText     =   "Quatidade total empenhada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox Txt_qtde_total_solicitada 
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
               TabIndex        =   56
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade total solicitada."
               Top             =   420
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. solicitada         Qtde. empenhada          Qtde. disponível"
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
               Left            =   9720
               TabIndex        =   123
               Top             =   210
               Width           =   5010
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   0
               TabIndex        =   122
               Top             =   0
               Width           =   75
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
               TabIndex        =   121
               Top             =   480
               Width           =   1965
            End
         End
         Begin VB.TextBox Txt_ID_PC 
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
            Left            =   2400
            TabIndex        =   117
            Top             =   6540
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Operação da lista"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   12840
            TabIndex        =   113
            Top             =   8190
            Width           =   2310
            Begin VB.ComboBox Cmb_opcao_lista_Item 
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
               ItemData        =   "frmCompras_Requisicao.frx":20000
               Left            =   180
               List            =   "frmCompras_Requisicao.frx":2000A
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   170
               Width           =   1965
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
            Left            =   660
            TabIndex        =   101
            Text            =   "0"
            Top             =   6540
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtidcarteira 
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
            Left            =   1650
            TabIndex        =   100
            Top             =   6540
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Framelista 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   4755
            Left            =   60
            TabIndex        =   85
            Top             =   330
            Width           =   15105
            Begin VB.OptionButton OptQS_com 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Qt. solic. com."
               DisabledPicture =   "frmCompras_Requisicao.frx":2001F
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
               Left            =   13545
               TabIndex        =   42
               Top             =   2190
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptQS_est 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Qt. solic. est."
               DisabledPicture =   "frmCompras_Requisicao.frx":269F61
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
               Left            =   11715
               TabIndex        =   40
               Top             =   2190
               Width           =   1275
            End
            Begin VB.TextBox txtQS_est 
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
               Left            =   11640
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade da unidade estoque solicitada"
               Top             =   2400
               Width           =   1425
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2760
               Picture         =   "frmCompras_Requisicao.frx":4B3EA3
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Visualizar arquivo."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox Txt_conta_contabil 
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
               Left            =   4680
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               ToolTipText     =   "Conta contábil."
               Top             =   3000
               Width           =   10245
            End
            Begin VB.CheckBox chkRemessa 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Remessa"
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
               Left            =   13980
               TabIndex        =   33
               Top             =   1058
               Width           =   945
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
               ItemData        =   "frmCompras_Requisicao.frx":4B4465
               Left            =   9495
               List            =   "frmCompras_Requisicao.frx":4B4467
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   38
               TabStop         =   0   'False
               ToolTipText     =   "Unidade comercial."
               Top             =   2400
               Width           =   855
            End
            Begin VB.CommandButton cmdfiltrar 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2100
               Picture         =   "frmCompras_Requisicao.frx":4B4469
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Filtrar por código interno."
               Top             =   390
               Width           =   315
            End
            Begin VB.ComboBox Cmb_prioridade 
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
               ItemData        =   "frmCompras_Requisicao.frx":4B4884
               Left            =   12675
               List            =   "frmCompras_Requisicao.frx":4B488E
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   32
               ToolTipText     =   "Prioridade."
               Top             =   990
               Width           =   1200
            End
            Begin VB.TextBox Txt_descricao_comercial 
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
               Locked          =   -1  'True
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               TabStop         =   0   'False
               ToolTipText     =   "Descrição comercial."
               Top             =   1620
               Width           =   9405
            End
            Begin VB.TextBox txtdetalheitem 
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
               TabIndex        =   45
               ToolTipText     =   "Detalhe."
               Top             =   3000
               Width           =   1965
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
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   180
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   36
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   2400
               Width           =   8415
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
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   9600
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               ToolTipText     =   "Observações."
               Top             =   1620
               Width           =   5325
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
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   990
               Width           =   10965
            End
            Begin VB.TextBox txtN_Estoque 
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
               TabIndex        =   21
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   1935
            End
            Begin VB.TextBox txtQS_com 
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
               Left            =   13500
               TabIndex        =   43
               ToolTipText     =   "Quantidade comercial solicitada."
               Top             =   2400
               Width           =   1425
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
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   10350
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade em estoque."
               Top             =   2400
               Width           =   1275
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
               ForeColor       =   &H00000000&
               Height          =   330
               ItemData        =   "frmCompras_Requisicao.frx":4B48A3
               Left            =   8610
               List            =   "frmCompras_Requisicao.frx":4B48A5
               Locked          =   -1  'True
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   37
               TabStop         =   0   'False
               ToolTipText     =   "Unidade de estoque."
               Top             =   2400
               Width           =   855
            End
            Begin VB.CommandButton cmdEscolher_item 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2430
               Picture         =   "frmCompras_Requisicao.frx":4B48A7
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Localizar produtos/serviços."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox cmbStatus 
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
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "Status."
               Top             =   390
               Width           =   5415
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
               Left            =   13080
               Picture         =   "frmCompras_Requisicao.frx":4B49A9
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Abrir calculadora para cálculo de peso."
               Top             =   2400
               Width           =   315
            End
            Begin VB.Frame Frame14 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto/serviço"
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
               Left            =   11580
               TabIndex        =   87
               Top             =   180
               Width           =   3345
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
                  TabIndex        =   29
                  Top             =   270
                  Width           =   1605
               End
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
                  TabIndex        =   28
                  Top             =   270
                  Width           =   1335
               End
            End
            Begin VB.TextBox txtOrdem 
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
               Left            =   2160
               TabIndex        =   46
               ToolTipText     =   "Número da ordem de produção."
               Top             =   3000
               Width           =   1095
            End
            Begin VB.ComboBox Cmb_OS 
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
               ItemData        =   "frmCompras_Requisicao.frx":4B4C12
               Left            =   3270
               List            =   "frmCompras_Requisicao.frx":4B4C14
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   47
               ToolTipText     =   "Número da OS."
               Top             =   3000
               Width           =   1395
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Empenhos da ordem de produção (Pedidos)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1365
               Left            =   0
               TabIndex        =   86
               Top             =   3390
               Width           =   15105
               Begin MSComctlLib.ListView Lista_pedidos 
                  Height          =   1005
                  Left            =   180
                  TabIndex        =   49
                  Top             =   210
                  Width           =   14745
                  _ExtentX        =   26009
                  _ExtentY        =   1773
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
                  NumItems        =   12
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Object.Tag             =   "N"
                     Text            =   "ID"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Tag             =   "N"
                     Text            =   "Cód. carteira"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Object.Tag             =   "N"
                     Text            =   "Ped. interno"
                     Object.Width           =   1676
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
                     Object.Width           =   6265
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   5
                     Object.Tag             =   "T"
                     Text            =   "Cód. interno"
                     Object.Width           =   1940
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
                     Text            =   "Cod. de ref."
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   8
                     Object.Tag             =   "T"
                     Text            =   "Descrição"
                     Object.Width           =   6265
                  EndProperty
                  BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   9
                     Object.Tag             =   "N"
                     Text            =   "Qtde. vend."
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   10
                     Object.Tag             =   "D"
                     Text            =   "Prazo final"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   11
                     Object.Tag             =   "T"
                     Text            =   "Tipo"
                     Object.Width           =   0
                  EndProperty
               End
            End
            Begin MSMask.MaskEdBox txtprazo 
               Height          =   315
               Left            =   11160
               TabIndex        =   31
               ToolTipText     =   "Prazo de entrega."
               Top             =   990
               Width           =   1125
               _ExtentX        =   1984
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
               ItemData        =   "frmCompras_Requisicao.frx":4B4C16
               Left            =   3180
               List            =   "frmCompras_Requisicao.frx":4B4C18
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   25
               ToolTipText     =   "Código de referência."
               Top             =   390
               Width           =   2865
            End
            Begin VB.TextBox txtReferencia 
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
               Left            =   3180
               MaxLength       =   50
               TabIndex        =   26
               ToolTipText     =   "Código de referência."
               Top             =   390
               Visible         =   0   'False
               Width           =   2865
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Conta contábil"
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
               Left            =   9285
               TabIndex        =   116
               Top             =   2790
               Width           =   1035
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
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
               Left            =   9600
               TabIndex        =   111
               Top             =   2190
               Width           =   645
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Prioridade*"
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
               Left            =   12915
               TabIndex        =   110
               Top             =   780
               Width           =   810
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descrição comercial"
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
               Left            =   4185
               TabIndex        =   105
               Top             =   1410
               Width           =   1395
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
               Left            =   8490
               TabIndex        =   99
               Top             =   180
               Width           =   555
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo entrega"
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
               Left            =   11205
               TabIndex        =   98
               Top             =   780
               Width           =   1020
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Detalhe"
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
               Left            =   885
               TabIndex        =   97
               Top             =   2790
               Width           =   555
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
               Left            =   11677
               TabIndex        =   96
               Top             =   1410
               Width           =   870
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. estoque"
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
               Left            =   10455
               TabIndex        =   95
               Top             =   2190
               Width           =   1050
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
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
               Index           =   0
               Left            =   8745
               TabIndex        =   94
               Top             =   2190
               Width           =   585
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
               Index           =   0
               Left            =   4117
               TabIndex        =   93
               Top             =   2190
               Width           =   540
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Index           =   1
               Left            =   5317
               TabIndex        =   92
               Top             =   780
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
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
               Index           =   0
               Left            =   465
               TabIndex        =   91
               Top             =   180
               Width           =   1350
            End
            Begin VB.Image imgCalendario 
               Height          =   360
               Left            =   12285
               Picture         =   "frmCompras_Requisicao.frx":4B4C1A
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   960
               Width           =   330
            End
            Begin VB.Label Label57 
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
               Left            =   3855
               TabIndex        =   90
               Top             =   180
               Width           =   1500
            End
            Begin VB.Label Label26 
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
               Left            =   2602
               TabIndex        =   89
               Top             =   2790
               Width           =   210
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "OS"
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
               Left            =   3855
               TabIndex        =   88
               Top             =   2790
               Width           =   210
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   840
            Left            =   -74940
            TabIndex        =   83
            Top             =   330
            Width           =   15105
            Begin VB.TextBox txtPercentualCentro 
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
               Left            =   13770
               MaxLength       =   50
               TabIndex        =   53
               ToolTipText     =   "Percentual."
               Top             =   390
               Width           =   1155
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
               Left            =   180
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   52
               ToolTipText     =   "Centro de custo."
               Top             =   390
               Width           =   13580
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Percentual*"
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
               Left            =   13920
               TabIndex        =   115
               Top             =   180
               Width           =   855
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Centro de custo*"
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
               Left            =   6348
               TabIndex        =   84
               Top             =   180
               Width           =   1245
            End
         End
         Begin VB.TextBox txtIDCentro 
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
            Left            =   -72300
            MouseIcon       =   "frmCompras_Requisicao.frx":4B509D
            MousePointer    =   99  'Custom
            TabIndex        =   82
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   4170
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComctlLib.ListView lista 
            Height          =   3075
            Left            =   60
            TabIndex        =   50
            Top             =   5100
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   5424
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
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "Nº do item"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Cód. int."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   4612
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
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Qt. sol. est."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Qt. sol. com."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Família"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "D"
               Text            =   "Prazo entr."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Tag             =   "T"
               Text            =   "Detalhe"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "Ordem"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Tag             =   "N"
               Text            =   "OS"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_custo 
            Height          =   7200
            Left            =   -74940
            TabIndex        =   54
            Top             =   1185
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   12700
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
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   22781
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "ID_CC"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_empenhos 
            Height          =   7185
            Left            =   -74940
            TabIndex        =   55
            Top             =   330
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   12674
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
         Begin DrawSuite2022.USProgressBar PBLista1 
            Height          =   255
            Left            =   30
            TabIndex        =   124
            Top             =   8395
            Width           =   12705
            _ExtentX        =   22410
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
      End
      Begin MSComctlLib.ListView lista_req 
         Height          =   5145
         Left            =   -74925
         TabIndex        =   12
         Top             =   3930
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   9075
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Nº solicitação"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3441
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Setor "
            Object.Width           =   3441
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Dt. aprov."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Responsável aprov."
            Object.Width           =   3441
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Setor"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Validada"
            Object.Width           =   1499
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCompras_Requisicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_solicitacao     As Boolean 'OK
Dim Novo_solicitacao1       As Boolean 'OK
Dim Novo_solicitacao1_Custo As Boolean 'OK
Public StrSql_solicitacao   As String 'OK
Dim TBLISTA_Solicitacao     As ADODB.Recordset 'OK
Dim NItem                   As Integer 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=i46JnPbSe98&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=36&feature=plcp")

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
    USMsgBox ("Informe o código interno do produto."), vbInformation, "CAPRIND v5.0"
    txtN_Estoque.Text = ""
    txtN_Estoque.SetFocus
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
With Txt_descricao_comercial
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
If chkAuto.Value = 1 Or chkManual.Value = 1 Then
    cmbRef.Visible = False
    txtreferencia.Visible = True
End If

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
With Txt_descricao_comercial
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
cmbRef.Visible = True
txtreferencia.Visible = False

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

If txtNumero = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_requisicao order by Requisicaotexto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Requisicaotexto = '" & txtNumero & "'")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimpaCampos
        txtNumero = TBLISTA!Requisicaotexto
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from Compras_requisicao where Requisicaotexto = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaReq
        ProcCarregaLista
    Else
        USMsgBox ("Fim dos cadastros de solicitação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_solicitacao1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Item_Click()
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With
ProcHabDesabBotoesProdServ

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabDesabBotoesProdServ()
On Error GoTo tratar_erro

With USToolBar2
    .ButtonState(2) = 0
    If Cmb_opcao_lista_Item = "Excluir" Then
        .ButtonState(3) = 0
        .ButtonState(7) = 5
    Else
        .ButtonState(3) = 5
        .ButtonState(7) = 0
    End If
    .Refresh
End With

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
        Case "Excluir"
            .ButtonState(4) = 0
            .ButtonState(8) = 5
            .ButtonState(10) = 5
            .ButtonState(11) = 5
        Case "Status"
            .ButtonState(4) = 5
            .ButtonState(8) = 0
            .ButtonState(10) = 5
            .ButtonState(11) = 5
        Case "Validação"
            .ButtonState(4) = 5
            .ButtonState(8) = 5
            .ButtonState(10) = 0
            .ButtonState(11) = 5
        Case "Aprovação"
            .ButtonState(4) = 5
            .ButtonState(8) = 5
            .ButtonState(10) = 5
            .ButtonState(11) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_un_com_Click()
On Error GoTo tratar_erro

If txtN_Estoque <> "" Then
    If txtQS_com <> "" Then
        If cmbun <> Cmb_un_com Then
            txtQS_est = FunFormataCasasDecimais(4, FunConversaoFinalUn(cmbun, Cmb_un_com, txtQS_com, txtN_Estoque, True))
        Else
            txtQS_est = FunFormataCasasDecimais(4, txtQS_com)
        End If
        
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
            txtQS_com_PC = FunCalculaQtdePC(txtN_Estoque, txtQS_com, True, Cmb_un_com)
        Else
            txtQS_com_PC = ""
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbun_Click()
On Error GoTo tratar_erro

With txtQS_com
    .Locked = False
    .TabStop = True
End With
If cmbun <> "" Then Cmb_un_com = cmbun

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_dados_cancelamento_Click()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_requisicao where Requisicaotexto = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Dados do cancelamento: " & vbCrLf & "Data: " & Format(TBAbrir!datacancelada, "dd/mm/yy") & " " & vbCrLf & "Responsável: " & TBAbrir!cancelou & " " & vbCrLf & "Motivo: " & TBAbrir!motivo), vbInformation, "CAPRIND v5.0"
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtN_Estoque = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtN_Estoque & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
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

If txtN_Estoque = "" Or OptQS_est.Value = False Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Desenho, Unidade, Un_Kg, peso_metro from projproduto where desenho = '" & txtN_Estoque & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Engenharia = False
    Compras_Requisicao = True
    Compras_Cotacao = False
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

Private Sub ProcCopiar()
On Error GoTo tratar_erro
Dim ContAntigo As Integer 'OK

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtNumero = "" Then
    Acao = "copiar"
    NomeCampo = "a solicitação"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_solicitacao = True Then
    USMsgBox ("Salve a solicitação antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar esta solicitação?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ContAntigo = Txt_ID_req
    Set TBSolicitacao = CreateObject("adodb.recordset")
    TBSolicitacao.Open "Select * from compras_requisicao", Conexao, adOpenKeyset, adLockOptimistic
    TBSolicitacao.AddNew
    TBSolicitacao!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBSolicitacao!solicitado = pubUsuario
    TBSolicitacao!Data_Solicitacao = Format(Date, "dd/mm/yy")
    TBSolicitacao!setorsolic = IIf(pubSetor = "", Null, pubSetor)
    TBSolicitacao!Observacao = IIf(txtobssolicitacao = "", Null, txtobssolicitacao)
    TBSolicitacao!status = "ABERTA"
    ProcCriarNovoNumero
    TBSolicitacao!Requisicaotexto = a
    TBSolicitacao.Update
    Txt_ID_req = TBSolicitacao!ID_Requisicao
    TBSolicitacao.Close
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_pedido_lista where ID_Requisicao = " & ContAntigo & " and Status_Item <> 'CANCELADO' order by IdLista", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from Compras_pedido_lista", Conexao, adOpenKeyset, adLockOptimistic
            TBGravar.AddNew
            TBGravar!ID_Requisicao = Txt_ID_req
            TBGravar!Codproduto = TBAbrir!Codproduto
            TBGravar!Tipo = TBAbrir!Tipo
            TBGravar!CODIGO = TBAbrir!CODIGO
            TBGravar!Status_Item = "REQUISIT."
            TBGravar!Un = TBAbrir!Un
            TBGravar!Unidade_com = TBAbrir!Unidade_com
            TBGravar!Familia = TBAbrir!Familia
            TBGravar!solicitado = pubUsuario
            TBGravar!setorsolic = pubSetor
            TBGravar!Descricao = TBAbrir!Descricao
            TBGravar!Descricao_comercial = TBAbrir!Descricao_comercial
            TBGravar!quant_req = TBAbrir!quant_req
            TBGravar!quant_req_PC = TBAbrir!quant_req_PC
            TBGravar!Desenho = TBAbrir!Desenho
            TBGravar!N_referencia = TBAbrir!N_referencia
            TBGravar!detalheitem = TBAbrir!detalheitem
            TBGravar!prazoreq = IIf(IsNull(TBAbrir!prazoreq), Null, Format(TBAbrir!prazoreq, "dd/mm/yy"))
            TBGravar!Prioridade = TBAbrir!Prioridade
            TBGravar!Remessa = TBAbrir!Remessa
            TBGravar!Obs = TBAbrir!Obs
            TBGravar!Ordem = TBAbrir!Ordem
            TBGravar!OS = TBAbrir!OS
            TBGravar!ID_PC = TBAbrir!ID_PC
            TBGravar.Update
            
            'Copiar centro de custo
            Set TBCarteira = CreateObject("adodb.recordset")
            TBCarteira.Open "Select * from Compras_pedido_lista_custo where IDLista = " & TBAbrir!IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                Do While TBCarteira.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Compras_pedido_lista_custo", Conexao, adOpenKeyset, adLockOptimistic
                    TBAbrir.AddNew
                    TBAbrir!ID_Requisicao = TBGravar!ID_Requisicao
                    TBAbrir!IDlista = TBGravar!IDlista
                    TBAbrir!ID_CC = TBCarteira!ID_CC
                    TBAbrir!Data = Date
                    TBAbrir!Responsavel = pubUsuario
                    TBAbrir.Update
                    TBCarteira.MoveNext
                Loop
            End If
            TBCarteira.Close
            
            TBGravar.Close
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_requisicao where id_requisicao = " & Txt_ID_req, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = False Then
        ProcLimpaCampos
        ProcAbrir
    End If
    ProcCarregaLista_Req (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista_req.ListItems.Count <> 0 Then
        Lista_req.SelectedItem = Lista_req.ListItems(CodigoLista)
        Lista_req.SetFocus
    End If
    USMsgBox ("Solicitação copiada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Outros/Solicitação"
    Evento = "Novo"
    ID_documento = Txt_ID_req
    Documento = "Nº solicitação: " & txtNumero
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcHabilitaCamposSolic
    Novo_solicitacao = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEscolher_item_Click()
On Error GoTo tratar_erro

If ProcVerifUsuario(True) = False Then
    Permitido = False
    Exit Sub
End If
If ProcVerifSatus("alterar este produto/serviço", True) = False Then Exit Sub
If cmbStatus <> "REQUISITADO" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois o mesmo está " & cmbStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Outros_solicitacaoPCP = False
frmcompras_Req_EscolherProduto.Show 1

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
If ProcVerifUsuario(True) = False Then Exit Sub
Select Case SSTab2.Tab
    Case 0: ProcNovoItem
    Case 1: ProcNovoItem_Custo
    Case 2: ProcNovoEmpenho
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoItem()
On Error GoTo tratar_erro

If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "produto/serviço", "criar novo", False, True) = False Then Exit Sub
If ProcVerifSatus("criar novo produto/serviço", True) = False Then Exit Sub
ProcLimpaCampos2
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select CODIGO from Compras_pedido_lista WHERE ID_Requisicao = " & Txt_ID_req & " order by codigo desc", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then NItem = TBCompras!CODIGO + 1 Else NItem = 1
TBCompras.Close
Novo_solicitacao1 = True
Framelista.Enabled = True
txtN_Estoque.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoItem_Custo()
On Error GoTo tratar_erro

If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "centro de custo", "criar novo", False, True) = False Then Exit Sub
If ProcVerifSatus("criar novo centro de custo", True) = False Then Exit Sub
'Verifica se o produto controla estoque e não permitie adicionar o centro de custo
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select codproduto from projproduto where desenho = '" & txtN_Estoque & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    USMsgBox ("Não é permitido criar centro de custo para este produto/serviço, pois o mesmo movimenta estoque."), vbExclamation, "CAPRIND v5.0"
    TBProduto.Close
    Exit Sub
End If
TBProduto.Close
ProcLimpaCamposCusto True, True
Frame8.Enabled = True
Cmb_centro.SetFocus
Novo_solicitacao1_Custo = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoEmpenho()
On Error GoTo tratar_erro

If ProcVerifSatus("criar novo empenho", True) = False Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select IDlista from Compras_pedido_lista where IDlista = " & TXTIDLista & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then Sit_REG = 0 Else Sit_REG = 1
TBAbrir.Close

Compras_Requisicao = True
Compras_Cotacao = False
Compras_Pedido = False
frmProd_Lista_Produto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtNumero = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_requisicao order by Requisicaotexto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Requisicaotexto = '" & txtNumero & "'")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimpaCampos
        txtNumero = TBLISTA!Requisicaotexto
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from Compras_requisicao where Requisicaotexto = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcCarregaReq
        ProcCarregaLista
    Else
        USMsgBox ("Fim dos cadastros de solicitação."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_solicitacao1 = False

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

chkRemessa.Enabled = True
If txtN_Estoque <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto WHERE desenho = '" & txtN_Estoque.Text & "' and DtValidacao IS NOT NULL and (Compras = 'True' or Producao = 'True')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        txtdescricao.Text = ""
        Txt_descricao_comercial = ""
        cmbun.ListIndex = -1
        Cmb_un_com.ListIndex = -1
        cmbfamilia.ListIndex = -1
        txtQE.Text = "0,0000"
        
        If TBProduto!Bloqueado = False Then
            txtN_Estoque = TBProduto!Desenho
            txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
            Txt_descricao_comercial = IIf(IsNull(TBProduto!descricaotecnica), "", TBProduto!descricaotecnica)
            If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then cmbun.Text = TBProduto!Unidade
            'If IsNull(TBProduto!Unidade) = False And TBProduto!Unidade <> "" Then Cmb_un_com.Text = TBProduto!Unidade
            If IsNull(TBProduto!Unidade_com) = False And TBProduto!Unidade_com <> "" Then Cmb_un_com.Text = TBProduto!Unidade_com
            If IsNull(TBProduto!Classe) = False And TBProduto!Classe <> "" Then cmbfamilia.Text = TBProduto!Classe
            If IsNull(TBProduto!ID_PC) = False And TBProduto!ID_PC <> "" Then ProcCarregaPC TBProduto!ID_PC
2:
            txtQE = Format(FunVerificaQtdeEstoque(txtN_Estoque, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
            If TBProduto!Instrumento = False Then ProcCarregaComboCodRef cmbRef, "P.Desenho = '" & txtN_Estoque & "'", 0, "", False, True
            
            With cmbun
                If TBProduto!Estoque = True Then
                    .Locked = True
                    .TabStop = False
                Else
                    .Locked = False
                    .TabStop = True
                End If
            End With
    
            With chkRemessa
                If TBProduto!Compras = False Then
                    .Value = 1
                    .Enabled = False
                Else
                    .Value = 0
                    .Enabled = True
                End If
            End With
        Else
            USMsgBox ("Não é permitido utilizar este " & IIf(TBProduto!Tipo <> "S", "produto", "serviço") & ", pois o mesmo está bloqueado."), vbExclamation, "CAPRIND v5.0"
        End If
        ProcBloqueiaCampos
    Else
        txtdescricao.Text = ""
        Txt_descricao_comercial = ""
        cmbun.ListIndex = -1
        Cmb_un_com.ListIndex = -1
        cmbfamilia.ListIndex = -1
        txtQE.Text = "0,0000"
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos
    End If
Else
    If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado a unidade ou familia desse registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 2
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Solicitacao.AbsolutePage <> 2 Then
    If TBLISTA_Solicitacao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Solicitacao.PageCount - 1)
    Else
        TBLISTA_Solicitacao.AbsolutePage = TBLISTA_Solicitacao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Solicitacao.AbsolutePage)
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
    TBLISTA_Solicitacao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Solicitacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Solicitacao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Solicitacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Solicitacao.AbsolutePage <> -3 Then
    If TBLISTA_Solicitacao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Solicitacao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Solicitacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Solicitacao.AbsolutePage = TBLISTA_Solicitacao.PageCount
ProcExibePagina (TBLISTA_Solicitacao.AbsolutePage)

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
            Case vbKeyF7: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF8: ProcCopiar
            Case vbKeyF9: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista_req, "Outros/Solicitação"
            Case vbKeyF10: If Cmb_opcao_lista = "Aprovação" Then ProcValidarRegistros Lista_req, "Outros/Solicitação/Autorizar solicitação"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo1
            Case vbKeyF3: If USToolBar2.ButtonState(2) = 0 Then procSalvar1
            Case vbKeyF4: If Cmb_opcao_lista_Item = "Excluir" Then ProcExcluir1
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: If Cmb_opcao_lista_Item = "Status" Then ProcStatus1
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Sub ProcAbrir()
On Error GoTo tratar_erro

ProcCarregaReq
ProcHabilitaCamposSolic
If TBCompras!status = "CANCELADA" Then
    Frame1.Enabled = False
Else
    Frame1.Enabled = True
    If TBCompras!status = "LIBERADA" Then
        With txtobssolicitacao
            .Locked = True
            .TabStop = False
        End With
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaReq()
On Error GoTo tratar_erro

If IsNull(TBCompras!ID_empresa) = False And TBCompras!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBCompras!ID_empresa
Txt_ID_req = TBCompras!ID_Requisicao
Caption = "Outros - Solicitação - (Solicitação : " & IIf(IsNull(TBCompras!Requisicaotexto), "", TBCompras!Requisicaotexto) & ")"
txtNumero = IIf(IsNull(TBCompras!Requisicaotexto), "", TBCompras!Requisicaotexto)
If TBCompras!status = "LIBERADA" Then
    txtAutorizado.Text = IIf(IsNull(TBCompras!Autorizado), "", (TBCompras!Autorizado))
    txtData_Autorizacao = IIf(IsNull(TBCompras!Data_autorizacao), "", TBCompras!Data_autorizacao)
End If
txtDtValidacao = IIf(IsNull(TBCompras!DtValidacao), "", (TBCompras!DtValidacao))
txtRespValidacao = IIf(IsNull(TBCompras!RespValidacao), "", TBCompras!RespValidacao)
txtSolicitado.Text = IIf(IsNull(TBCompras!solicitado), "", (TBCompras!solicitado))
txtData_Solicitacao.Text = IIf(IsNull(TBCompras!Data_Solicitacao), "", (Format(TBCompras!Data_Solicitacao, "dd/mm/yy")))
txtSetor_Solicitado.Text = IIf(IsNull(TBCompras!setorsolic), "", (TBCompras!setorsolic))
txtobssolicitacao.Text = IIf(IsNull(TBCompras!Observacao), "", TBCompras!Observacao)
txtStatus.Text = IIf(IsNull(TBCompras!status), "", TBCompras!status)
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Compras_pedido_lista Where ID_Requisicao = " & Txt_ID_req & " order by idlista desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!IDlista
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Unidade_com), "", TBLISTA!Unidade_com)
            If TBLISTA!Un <> TBLISTA!Unidade_com Then valor = FunConversaoFinalUn(TBLISTA!Un, TBLISTA!Unidade_com, TBLISTA!quant_req, TBLISTA!Desenho, True) Else valor = TBLISTA!quant_req
            .Item(.Count).SubItems(6) = FunFormataCasasDecimais(4, valor)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!quant_req), "", FunFormataCasasDecimais(4, TBLISTA!quant_req))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Familia), "", TBLISTA!Familia)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!prazoreq), "", Format(TBLISTA!prazoreq, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!detalheitem), "", TBLISTA!detalheitem)
            If IsNull(TBLISTA!Status_Item) = False Then
                If TBLISTA!Status_Item = "REQUISIT." Then
                    .Item(.Count).SubItems(11) = "REQUISITADO"
                ElseIf TBLISTA!Status_Item = "N_RECEBIDO" Then
                        .Item(.Count).SubItems(11) = "COMPRADO"
                    ElseIf TBLISTA!Status_Item = "PARCIAL" Then
                        .Item(.Count).SubItems(11) = "RECEBIDO PARCIAL"
                    Else
                        .Item(.Count).SubItems(11) = TBLISTA!Status_Item
                End If
            End If
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!OS), "", TBLISTA!OS)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Custo()
On Error GoTo tratar_erro

Lista_custo.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CPLC.*, US.Codigo, US.Setor from Compras_pedido_lista_custo CPLC INNER JOIN Usuarios_setor US ON CPLC.ID_CC = US.ID where CPLC.ID_requisicao = " & Txt_ID_req & " and CPLC.idlista = " & TXTIDLista & " order by US.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista_custo.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            .Item(.Count).SubItems(3) = TBLISTA!ID_CC
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos2()
On Error GoTo tratar_erro

TXTIDLista = 0
txtN_Estoque.Text = ""
chkAuto.Value = 0
chkManual.Value = 0
cmbfamilia.ListIndex = -1
cmbun.ListIndex = -1
Cmb_un_com.ListIndex = -1
txtQE.Text = ""
txtQS_est = ""
txtQS_com.Text = ""
txtQS_com_PC = ""
txtdescricao.Text = ""
Txt_descricao_comercial = ""
txtdetalheitem.Text = ""
txtprazo.Text = "__/__/____"
Cmb_prioridade.ListIndex = -1
chkRemessa.Value = 0
txtObs = ""
cmbStatus = "REQUISITADO"
txtOrdem = ""
With Cmb_OS
    .ListIndex = -1
    .Locked = True
    .TabStop = False
End With
Txt_conta_contabil = ""
cmbRef.Clear
txtreferencia = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposCusto(VerifCCProd As Boolean, VerifRespCC As Boolean)
On Error GoTo tratar_erro

txtIDCentro = 0
txtPercentualCentro = ""
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select A.* from Acessos A INNER JOIN Usuarios U ON A.IDUsuario = U.IDUsuario where U.Usuario = '" & txtSolicitado & "' and A.Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    ProcCarregaComboSetor Cmb_centro, "ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and setor is not null and (Consolidacao = 'False' or Consolidacao is null)", txtN_Estoque, VerifCCProd, True, VerifRespCC, "", True, False
Else
    ProcCarregaComboSetor Cmb_centro, "US.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), txtN_Estoque, VerifCCProd, True, VerifRespCC, txtSolicitado, True, False
End If
TBAcessos.Close
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID_req = 0
txtNumero.Text = ""
txtStatus.Text = "ABERTA"
txtSolicitado.Text = pubUsuario
txtSetor_Solicitado.Text = pubSetor
txtData_Solicitacao.Text = Format(Date, "dd/mm/yy")
txtData_Autorizacao.Text = ""
txtAutorizado.Text = ""
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtobssolicitacao.Text = ""
CodigoLista = 0
Caption = "Outros - Solicitação"
ProcHabilitaCamposSolic

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposPedido()
On Error GoTo tratar_erro

txtOrdem = ""
Cmb_OS.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Procenviadadoslista()
On Error GoTo tratar_erro

TBCompras!ID_Requisicao = Txt_ID_req
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto, Tipo from Projproduto where Desenho = '" & txtN_Estoque & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    TBCompras!Codproduto = TBProduto!Codproduto
    If TBProduto!Tipo = "S" Then TBCompras!Tipo = "S" Else TBCompras!Tipo = "P"
End If
TBProduto.Close
TBCompras!CODIGO = NItem
If cmbStatus = "REQUISITADO" Then TBCompras!Status_Item = "REQUISIT."
TBCompras!Un = cmbun.Text
TBCompras!Unidade_com = Cmb_un_com.Text
TBCompras!Familia = cmbfamilia.Text
TBCompras!solicitado = txtSolicitado.Text
TBCompras!setorsolic = txtSetor_Solicitado.Text
TBCompras!Descricao = txtdescricao.Text
TBCompras!Descricao_comercial = Txt_descricao_comercial
TBCompras!quant_req = txtQS_com.Text
TBCompras!quant_req_PC = IIf(txtQS_com_PC = "", Null, txtQS_com_PC)
TBCompras!Desenho = txtN_Estoque.Text
TBCompras!N_referencia = IIf(cmbRef.Text = "", Null, cmbRef.Text)
TBCompras!detalheitem = txtdetalheitem.Text
If txtprazo.Text <> "__/__/____" Then TBCompras!prazoreq = txtprazo.Text Else TBCompras!prazoreq = Null
TBCompras!Prioridade = Cmb_prioridade
If chkRemessa.Value = 1 Then TBCompras!Remessa = True Else TBCompras!Remessa = False
TBCompras!Obs = txtObs

TBCompras!Ordem = IIf(txtOrdem = "", Null, txtOrdem)
TBCompras!OS = IIf(Cmb_OS = "", Null, Cmb_OS)
If Cmb_OS <> "" Then Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where IDlista = " & TXTIDLista

TBCompras!ID_PC = IIf(Txt_ID_PC = "", Null, Txt_ID_PC)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcHabilitaCamposSolic()
On Error GoTo tratar_erro

Frame1.Enabled = True
With txtobssolicitacao
    .Locked = False
    .TabStop = True
End With

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
If ProcVerifUsuario(True) = False Then Exit Sub
Select Case SSTab2.Tab
    Case 0: ProcSalvarItem
    Case 1: ProcSalvarItem_custo
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarItem()
On Error GoTo tratar_erro

If Framelista.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "o produto/serviço", "alterar", False, True) = False Then Exit Sub
If ProcVerifSatus("alterar este produto/serviço", True) = False Then Exit Sub
If cmbStatus <> "REQUISITADO" Then
    USMsgBox ("Não é permitido alterar este produto/serviço, pois o mesmo está " & cmbStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If chkAuto.Value = 0 And txtN_Estoque.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    txtN_Estoque.SetFocus
    Exit Sub
End If
If txtdescricao.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtdescricao.SetFocus
    Exit Sub
End If
If txtprazo <> "__/__/____" Then
    If IsDate(txtprazo) = False Then
        USMsgBox ("A data foi digitada incorretamente."), vbExclamation, "CAPRIND v5.0"
        txtprazo.SetFocus
        Exit Sub
    End If
End If
If Cmb_prioridade = "" Then
    NomeCampo = "a prioridade"
    ProcVerificaAcao
    Cmb_prioridade.SetFocus
    Exit Sub
End If
If Txt_descricao_comercial = "" Then
    NomeCampo = "a descrição comercial"
    ProcVerificaAcao
    Txt_descricao_comercial.SetFocus
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
Qtd = IIf(txtQS_com = "", 0, txtQS_com)
If Qtd <= 0 Then
    NomeCampo = "a quantidade solicitada"
    ProcVerificaAcao
    txtQS_com.SetFocus
    Exit Sub
End If
If txtOrdem <> "" And txtOrdem <> "0" Then
    If FunVerifOPCarregaOS(Cmb_OS, txtOrdem, True, False) = False Then
        txtOrdem.SetFocus
        Exit Sub
    End If
End If

If Novo_solicitacao1 = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from compras_pedido_lista where desenho = '" & txtN_Estoque.Text & "' And Status_item = 'REQUISIT.'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If USMsgBox("Já existe uma solicitação em aberto para este produto/serviço, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            TBProduto.Close
            Exit Sub
        End If
    End If
    TBProduto.Close
End If
If chkAuto.Value = 1 Then
    ProcNovoProdutoAuto
    If txtreferencia <> "" Then
        cmbRef.AddItem txtreferencia
        cmbRef = txtreferencia
    End If
    chkAuto.Value = 0
End If
If chkManual.Value = 1 Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Codproduto from projproduto where desenho = '" & txtN_Estoque.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        USMsgBox ("Já existe um produto/serviço cadastrado com este código interno, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtN_Estoque.SetFocus
        Exit Sub
    End If
    TBProduto.Close
    ProcNovoProdutoManual
    If txtreferencia <> "" Then
        cmbRef.AddItem txtreferencia
        cmbRef = txtreferencia
    End If
    chkManual.Value = 0
End If

'Verifica se o produto está cadastrado
If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where desenho = '" & txtN_Estoque & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then
        USMsgBox ("Não é permitido salvar este produto/serviço, pois o mesmo não está cadastrado."), vbExclamation, "CAPRIND v5.0"
        TBProduto.Close
        Exit Sub
    End If
    TBProduto.Close
End If

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_pedido_lista WHERE IDLista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then TBCompras.AddNew
Procenviadadoslista
TBCompras.Update
TXTIDLista = TBCompras!IDlista
TBCompras.Close
ProcCarregaLista
If Novo_solicitacao1 = True Then
    USMsgBox ("Novo produto/serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo produto/serviço"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar produto/serviço"
    If CodigoLista1 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista1)
        Lista.SetFocus
    End If
End If
Novo_solicitacao1 = False
'==================================
Modulo = "Outros/Solicitação"
ID_documento = TXTIDLista
Documento = "Nº solicitação: " & txtNumero
Documento1 = "Cód. interno: " & txtN_Estoque
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarItem_custo()
On Error GoTo tratar_erro

If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "o centro de custo", "alterar", False, True) = False Then Exit Sub
If ProcVerifSatus("alterar este centro de custo", True) = False Then Exit Sub
If cmbStatus <> "REQUISITADO" Then
    USMsgBox ("Não é permitido alterar este centro de custo, pois o mesmo está " & cmbStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If Cmb_centro = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_centro.SetFocus
    Exit Sub
End If

Permitido = False
ID_CC = 0
'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If Novo_solicitacao1_Custo = False Then ID_CC = Lista_custo.SelectedItem.ListSubItems(3)

    If ID_CC <> Cmb_centro.ItemData(Cmb_centro.ListIndex) Then
        Formulario = "Compras/Autorização de centro de custo sem previsão"
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select ID_PC from projproduto where desenho = '" & txtN_Estoque & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
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
TBTempo.Close

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_pedido_lista where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Novo_solicitacao1_Custo = True Then TextoFiltro = "IDlista = " & TBAbrir!IDlista Else TextoFiltro = "IDlista = " & TBAbrir!IDlista & " and ID_CC = " & Lista_custo.SelectedItem.ListSubItems(3)
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from Compras_pedido_lista_custo where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar!ID_Requisicao = Txt_ID_req
        TBGravar!IDlista = TBAbrir!IDlista
        TBGravar!Responsavel = pubUsuario
        TBGravar!Data = Date
        Evento = "Novo centro de custo"
    Else
        Evento = "Alterar centro de custo"
    End If
    TBGravar!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex)
    TBGravar!Percentual = IIf(txtPercentualCentro = "", Null, txtPercentualCentro)
    TBGravar.Update
    TBGravar.Close
    
    '==================================
    Modulo = "Outros/Solicitação"
    ID_documento = TBAbrir!IDlista
    Documento = "Nº solicitação: " & txtNumero
    Documento1 = "Cód. interno: " & TBAbrir!Desenho & " - Centro de custo: " & Cmb_centro
    ProcGravaEvento
    '==================================
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID from Compras_pedido_lista_custo where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtIDCentro = TBAbrir!ID
    End If
    TBAbrir.Close
    
    ProcCarregaLista_Custo
    If Novo_solicitacao1_Custo = True Then
        USMsgBox ("Novo centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        If Lista_custo.ListItems.Count <> 0 And CodigoLista2 <> 0 Then
            Lista_custo.SelectedItem = Lista_custo.ListItems(CodigoLista2)
            Lista_custo.SetFocus
        End If
    End If
    Novo_solicitacao1_Custo = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCopiar_CC()
On Error GoTo tratar_erro

If Novo_solicitacao1_Custo = True Then
    USMsgBox ("Informe o centro de custo na lista antes de copiar."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "o centro de custo", "copiar", False, True) = False Then Exit Sub
If ProcVerifSatus("copiar este centro de custo", True) = False Then Exit Sub
If cmbStatus <> "REQUISITADO" Then
    USMsgBox ("Não é permitido copiar este centro de custo, pois o mesmo está " & cmbStatus & "."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If Cmb_centro = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_centro.SetFocus
    Exit Sub
End If

Permitido1 = False
If USMsgBox("Deseja copiar o centro de custo para todos produtos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Compras_pedido_lista where ID_Requisicao = " & Txt_ID_req & " and Status_item = 'REQUISIT.' order by idlista desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
            If Permitido1 = False Then
                Permitido = False
                Set TBTempo = CreateObject("adodb.recordset")
                TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBTempo.EOF = False Then
                    Formulario = "Compras/Autorização de centro de custo sem previsão"
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select ID_PC from projproduto where desenho = '" & IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho) & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = True Then
                        If USMsgBox("Existe(m) produto(s) sem conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
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
                TBTempo.Close
            End If
        
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "select * from Compras_pedido_lista_custo where IDlista = " & TBAbrir!IDlista & " and ID_CC = " & Lista_custo.SelectedItem.ListSubItems(3), Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                TBGravar!ID_Requisicao = Txt_ID_req
                TBGravar!IDlista = TBAbrir!IDlista
                TBGravar!Responsavel = pubUsuario
                TBGravar!Data = Date
                Evento = "Novo centro de custo"
            Else
                Evento = "Alterar centro de custo"
            End If
            TBGravar!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex)
            TBGravar!Percentual = IIf(txtPercentualCentro = "", Null, txtPercentualCentro)
            TBGravar.Update
            TBGravar.Close
            
            '==================================
            Modulo = "Outros/Solicitação"
            ID_documento = TBAbrir!IDlista
            Documento = "Nº solicitação: " & txtNumero
            Documento1 = "Cód. interno: " & TBAbrir!Desenho & " - Centro de custo: " & Cmb_centro
            ProcGravaEvento
            '==================================
        
            TBAbrir.MoveNext
        Loop
        
        ProcCarregaLista_Custo
        USMsgBox ("Centro de custo copiado com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoAuto()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtN_Estoque = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtN_Estoque, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtN_Estoque = FunCriaNovoProdServ(False, "codmanual = 'False' and (subtipoitem = 0 or subtipoitem = 1 or subtipoitem = 4 or subtipoitem = 5)", txtN_Estoque, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir1()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0: ProcExcluirLista
    Case 1: ProcExcluirLista_Custo
    Case 2: ProcExcluirLista_Empenho
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirLista()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_pedido_lista where IDlista = " & .ListItems.Item(InitFor)
            Conexao.Execute "DELETE from Compras_pedido_lista_custo where IDlista = " & .ListItems.Item(InitFor)
            Conexao.Execute "DELETE from Compras_pedido_lista_empenhos where IDlista = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "Outros/Solicitação"
            Evento = "Excluir produto/serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº solicitação: " & txtNumero
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos2
    ProcLimpaCamposPedido
    ProcCarregaLista
    Framelista.Enabled = False
    Novo_solicitacao1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirLista_Custo()
On Error GoTo tratar_erro

Permitido = False
With Lista_custo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) centro(s) de custo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_pedido_lista_custo WHERE ID = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "Outros/Solicitação"
            Evento = "Excluir centro de custo"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº solicitação: " & txtNumero
            Documento1 = "Cód. interno: " & txtN_Estoque & " - Centro de custo: " & .ListItems.Item(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centro(s) de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Centro(s) de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCusto True, True
    Frame8.Enabled = False
    ProcCarregaLista_Custo
    Novo_solicitacao1_Custo = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirLista_Empenho()
On Error GoTo tratar_erro

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
            Modulo = "Outros/Solicitação"
            Evento = "Excluir empenho do produto/serviço"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Nº solicitação: " & txtNumero & " - Cód. interno: " & txtN_Estoque
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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 16, True
ProcCarregaToolBar2 Me, 15195, 11, True

Formulario = "Outros/Solicitação"
Direitos
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Validação"
Cmb_opcao_lista_Item = "Excluir"
SSTab1.Tab = 0
SSTab2.Tab = 0
ProcCarregaFamiliaUN
ProcCarregaComboEmpresa Cmb_empresa, False

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Outros/Solicitação"
Direitos
ProcCarregaFamiliaUN
ProcLimpaVariaveisPrincipais

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
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Compras_pedido_lista where IDlista = " & TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
    If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then cmbun = TBAbrir!Un
    If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaFamiliaUN()
On Error GoTo tratar_erro

ProcCarregaComboFamilia cmbfamilia, "Familia <> 'Null' and (Compras = 'True' or Fabricacao = 'True')", False
ProcCarregaComboUnidade cmbun, False
ProcCarregaComboUnidade Cmb_un_com, False
ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362S" Then
    If USMsgBox("Deseja realmente atualizar o número e o status das solicitações de compra?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBCompras = CreateObject("adodb.recordset")
        TBCompras.Open "Select * from Compras_requisicao order by ID_Requisicao", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras.EOF = False Then
            TBCompras.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBCompras.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBCompras.MoveFirst
            Do While TBCompras.EOF = False
                Ano = Right(Year(TBCompras!Data_Solicitacao), 2)
                If Right(TBCompras!Requisicaotexto, 3) <> "/" & Ano Then
                    IDAntigo = Right(TBCompras!Requisicaotexto, 5)
                    Conexao.Execute "Update Compras_pedido_lista Set ID_requisicao = " & TBCompras!ID_Requisicao & " where ID_requisicao = " & IDAntigo
                    
                    RequisicaoNovo = TBCompras!Requisicaotexto & "/" & Ano
                    Conexao.Execute "Update Manutencao Set Solicitacao = '" & RequisicaoNovo & "' where Solicitacao = '" & TBCompras!Requisicaotexto & "'"
                    TBCompras!Requisicaotexto = RequisicaoNovo
                End If
                
                If TBCompras!Autorizado = "" Or IsNull(TBCompras!Autorizado) = True Then
                    If TBCompras!status <> "CANCELADA" Or IsNull(TBCompras!status) = True Then TBCompras!status = "ABERTA"
                Else
                    If TBCompras!status <> "CANCELADA" Or IsNull(TBCompras!status) = True Then TBCompras!status = "LIBERADA"
                End If
                
                TBCompras.Update
                TBCompras.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBCompras.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Outros/Solicitação"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
        ProcCarregaLista_Req (1)
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmCompras_Requisicao_abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgCalendario_Click()
On Error GoTo tratar_erro

Faturamento = False
Compras_Pedido = False
Compras_Requisicao = True
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
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_req
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) solicitação(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Compras_requisicao where ID_Requisicao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from Compras_pedido_lista where ID_Requisicao = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from CPLE from Compras_pedido_lista_empenhos CPLE INNER JOIN Compras_pedido_lista CPL ON CPL.IDlista = CPLE.IDlista where CPL.ID_Requisicao = " & .ListItems(InitFor)
            
            'Manutenção
            Conexao.Execute "Update Manutencao_data Set Solicitacao = NULL where Solicitacao = '" & .ListItems(InitFor).SubItems(2) & "'"
            
            '==================================
            Modulo = "Outros/Solicitação"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº solicitação: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Solicitação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcLimparTudo
    ProcCarregaLista_Req (1)
    Novo_solicitacao = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista_req
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o status desta(s) solicitação(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBCotacao = CreateObject("adodb.recordset")
            TBCotacao.Open "Select * from Compras_requisicao where ID_Requisicao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCotacao.EOF = False Then
                If TBCotacao!status = "CANCELADA" Then
                    Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'REQUISIT.' where ID_Requisicao = " & .ListItems(InitFor)
                    TBCotacao!status = "ABERTA"
                    TBCotacao!cancelou = ""
                    TBCotacao!datacancelada = Null
                    TBCotacao!motivo = ""
                    TBCotacao.Update
                    '==================================
                    Modulo = "Outros/Solicitação"
                    Evento = "Alterar status"
                    ID_documento = .ListItems(InitFor)
                    Documento = "Nº solicitação: " & .ListItems(InitFor).SubItems(2)
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                Else
                    IDlista = .ListItems(InitFor)
                    Familiatext = .ListItems(InitFor).SubItems(2)
                    Outros_solicitacaoPCP = False
                    frmCompras_Requisicao_cancelar.Show 1
                End If
                TBCotacao.Close
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) solicitação(ões) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Solicitação(ões) alterada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista_Req (1)
    Novo_solicitacao = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus1()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o status deste(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from Compras_pedido_lista WHERE IDLista = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                If IsNull(TBCompras!Status_Item) = False Then
                    If TBCompras!Status_Item = "REQUISIT." Or TBCompras!Status_Item = "COTANDO" Then TBCompras!Status_Item = "CANCELADO" Else TBCompras!Status_Item = "REQUISIT."
                    TBCompras.Update
                    
                    '==================================
                    Modulo = "Outros/Solicitação"
                    Evento = "Alterar status"
                    ID_documento = .ListItems(InitFor)
                    Documento = "Nº solicitação: " & txtNumero
                    Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(2)
                    ProcGravaEvento
                    '==================================
                End If
            End If
            TBCompras.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("produto(s)/serviço(s) alterada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos2
    ProcLimpaCamposPedido
    ProcCarregaLista
    Framelista.Enabled = False
    Novo_solicitacao1 = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcImprimir()
On Error GoTo tratar_erro

If txtNumero.Text <> "" Then
    NomeRel = "Compras_requisicao.rpt"
    ProcImprimirRel "{Compras_requisicao.ID_Requisicao}= " & Txt_ID_req & " and {Compras_pedido_lista.Status_Item} <> 'CANCELADO'", ""
Else
    USMsgBox ("Informe a solicitação antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
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
Frame1.Enabled = True
txtobssolicitacao.SetFocus
Novo_solicitacao = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Framelista.Enabled = False
ProcLimpaCampos2
ProcLimpaCamposCusto True, True
Lista.ListItems.Clear
Lista_pedidos.ListItems.Clear
Lista_custo.ListItems.Clear
SSTab2.Tab = 0
Novo_solicitacao1 = False
Novo_solicitacao1_Custo = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_solicitacao = True Then
    If USMsgBox("A solicitação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_solicitacao = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_solicitacao1 = True Then
    If USMsgBox("O produto/serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvar1
        If Novo_solicitacao1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_solicitacao1_Custo = True Then
    If USMsgBox("O centro de custo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarItem_custo
        If Novo_solicitacao1_Custo = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_solicitacao = False
Novo_solicitacao1 = False
Novo_solicitacao1_Custo = False
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
If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_requisicao where ID_Requisicao = " & Txt_ID_req, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then
    TBCompras.AddNew
    TBCompras!status = "ABERTA"
    ProcCriarNovoNumero
    TBCompras!Requisicaotexto = a
Else
    If ProcVerifUsuario(True) = False Then Exit Sub
    If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "mesma", "a solicitação", "alterar", False, True) = False Then Exit Sub
    If ProcVerifSatus("alterar esta solicitação", True) = False Then Exit Sub
End If
TBCompras!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBCompras!Data_Solicitacao = IIf(txtData_Solicitacao = "", Date, txtData_Solicitacao)
TBCompras!solicitado = IIf(txtSolicitado = "", pubUsuario, txtSolicitado)
TBCompras!setorsolic = IIf(txtSetor_Solicitado = "", IIf(pubSetor = "", Null, pubSetor), txtSetor_Solicitado)
TBCompras!Observacao = IIf(txtobssolicitacao = "", Null, txtobssolicitacao)
TBCompras.Update
ProcAbrir
TBCompras.Close

If Novo_solicitacao = True Then
    USMsgBox ("Nova solicitação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_solicitacao = "Select * from Compras_requisicao where Requisicaotexto = '" & txtNumero & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    ProcCarregaLista_Req (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista_Req (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista_req.ListItems.Count <> 0 Then
        Lista_req.SelectedItem = Lista_req.ListItems(CodigoLista)
        Lista_req.SetFocus
    End If
End If
'==================================
Modulo = "Outros/Solicitação"
ID_documento = Txt_ID_req
Documento = "Nº solicitação: " & txtNumero
Documento1 = ""
ProcGravaEvento
'==================================
Novo_solicitacao = False

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
                If Cmb_opcao_lista_Item = "Excluir" Then
                    If ProcVerifUsuario(False) = False Then GoTo Proximo
                    If FunVerificaRegistroValidadoSemMsg("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, True) = False Then GoTo Proximo
                    If ProcVerifSatus("", False) = False Then GoTo Proximo
                    If .ListItems.Item(InitFor).SubItems(11) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(11) <> "CANCELADO" Then GoTo Proximo
                Else
                    If .ListItems.Item(InitFor).SubItems(11) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(11) <> "COTANDO" And .ListItems.Item(InitFor).SubItems(11) <> "CANCELADO" Then GoTo Proximo
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

Private Sub Lista_custo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_custo
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If ProcVerifUsuario(False) = False Then GoTo Proximo
                If FunVerificaRegistroValidadoSemMsg("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, True) = False Then GoTo Proximo
                If ProcVerifSatus("", False) = False Then GoTo Proximo
                If Lista.SelectedItem.ListSubItems(11) <> "REQUISITADO" And Lista.SelectedItem.ListSubItems(11) <> "CANCELADO" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_custo, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_custo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_custo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If ProcVerifUsuario(True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "este centro de custo", "excluir", False, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If ProcVerifSatus("excluir este centro de custo", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If Lista.SelectedItem.ListSubItems(11) <> "REQUISITADO" And Lista.SelectedItem.ListSubItems(11) <> "CANCELADO" Then
                USMsgBox ("Não é permitido excluir este centro de custo, pois o produto/serviço está " & Lista.SelectedItem.ListSubItems(11) & "."), vbExclamation, "CAPRIND v5.0"
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

Private Sub Lista_custo_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_custo.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CPLC.*, US.Codigo, US.Setor from Compras_pedido_lista_custo CPLC INNER JOIN Usuarios_setor US ON CPLC.ID_CC = US.ID where CPLC.id = " & Lista_custo.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposCusto False, False
    txtIDCentro = TBAbrir!ID
    
    If IsNull(TBAbrir!CODIGO) = False And TBAbrir!CODIGO <> "" Then
        Cmb_centro = TBAbrir!CODIGO & " - " & IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
    Else
        Cmb_centro = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
    End If
    txtPercentualCentro = IIf(IsNull(TBAbrir!Percentual), "", Format(TBAbrir!Percentual, "###,##0.0000000000"))
    
    CodigoLista2 = Lista_custo.SelectedItem.index
End If
TBAbrir.Close
Frame8.Enabled = True
Novo_solicitacao1_Custo = False

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
                If ProcVerifUsuario(False) = False Then GoTo Proximo
                If ProcVerifSatus("", False) = False Then GoTo Proximo
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
            qtde_solicitada = txtQS_com
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
            Qtd = txtQS_com
            valor = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Qtde_empenho) as Valor from Compras_pedido_lista_empenhos where IDcarteira = " & .SelectedItem.ListSubItems(1) & " and ID <> " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            If Qtd < (valor + Qtde) Then
                USMsgBox ("A quantidade empenhada não pode ser maior que a quantidade solicitada, favor alterar."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
            End If
                        
            NovoValor = Replace(Qtde, ",", ".")
            Conexao.Execute "Update Compras_pedido_lista_empenhos Set Qtde_empenho = " & NovoValor & " where ID = " & .SelectedItem
            
            USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Outros/Solicitação"
            Evento = "Empenhar produto/serviço"
            ID_documento = .SelectedItem
            Documento = "Nº solicitação: " & txtNumero & " - Cód. interno: " & txtN_Estoque
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
            If ProcVerifUsuario(True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If ProcVerifSatus("excluir este empenho", True) = False Then
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

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista_Item = "Excluir" Then
                If ProcVerifUsuario(True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If FunVerificaRegistroValidado("Compras_requisicao", "ID_Requisicao = " & Txt_ID_req, "solicitação", "este produto/serviço", "excluir", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If ProcVerifSatus("excluir este produto/serviço", True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                If .ListItems.Item(InitFor).SubItems(11) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(11) <> "CANCELADO" Then
                    USMsgBox ("Não é permitido excluir este produto/serviço, pois o mesmo está " & .ListItems.Item(InitFor).SubItems(11) & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
            Else
                If .ListItems.Item(InitFor).SubItems(11) <> "REQUISITADO" And .ListItems.Item(InitFor).SubItems(11) <> "COTANDO" And .ListItems.Item(InitFor).SubItems(11) <> "CANCELADO" Then
                    USMsgBox ("Não é permitido alterar o status deste produto/serviço, pois o mesmo está " & .ListItems.Item(InitFor).SubItems(11) & "."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
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
ProcLimpaCampos2
txtN_Estoque.Locked = False
txtN_Estoque.TabStop = True
TXTIDLista = Lista.SelectedItem
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from compras_pedido_lista where idlista = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtObs.Text = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
    txtN_Estoque.Text = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
    ProcCarregaComboCodRef cmbRef, "P.Desenho = '" & txtN_Estoque & "'", 0, "", False, True
    If IsNull(TBAbrir!N_referencia) = False And TBAbrir!N_referencia <> "" Then cmbRef = TBAbrir!N_referencia Else cmbRef.ListIndex = -1
1:
    txtdescricao.Text = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    Txt_descricao_comercial = IIf(IsNull(TBAbrir!Descricao_comercial), "", TBAbrir!Descricao_comercial)
    txtprazo.Text = IIf(IsNull(TBAbrir!prazoreq), "__/__/____", Format(TBAbrir!prazoreq, "dd/mm/yyyy"))
    If IsNull(TBAbrir!Prioridade) = False And TBAbrir!Prioridade <> "" Then Cmb_prioridade = TBAbrir!Prioridade
    If TBAbrir!Remessa = True Then chkRemessa.Value = 1 Else chkRemessa.Value = 0
    If IsNull(TBAbrir!Familia) = False And TBAbrir!Familia <> "" Then cmbfamilia = TBAbrir!Familia
    If IsNull(TBAbrir!Un) = False And TBAbrir!Un <> "" Then cmbun = TBAbrir!Un
    If IsNull(TBAbrir!Unidade_com) = False And TBAbrir!Unidade_com <> "" Then Cmb_un_com = TBAbrir!Unidade_com
    txtdetalheitem.Text = IIf(IsNull(TBAbrir!detalheitem), "", TBAbrir!detalheitem)
    txtQE = Format(FunVerificaQtdeEstoque(TBAbrir!Desenho, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""), "###,##0.0000")
    
    txtQS_com.Text = IIf(IsNull(TBAbrir!quant_req), "", Format(TBAbrir!quant_req, "###,##0.0000"))
    Txt_qtde_total_solicitada = txtQS_com
    
    txtQS_com_PC = IIf(IsNull(TBAbrir!quant_req_PC), "", TBAbrir!quant_req_PC)
    
    If IsNull(TBAbrir!Status_Item) = False Then
        cmbStatus.Text = Lista.SelectedItem.ListSubItems(11)
        If cmbStatus <> "REQUISITADO" Then Procbloqueia Else ProcDesbloqueia
    Else
        ProcDesbloqueia
    End If
    
    txtOrdem = IIf(IsNull(TBAbrir!Ordem), "", TBAbrir!Ordem)
    If txtOrdem <> "" And txtOrdem <> "0" Then
        If FunVerifOPCarregaOS(Cmb_OS, txtOrdem, False, True) = True Then
            If IsNull(TBAbrir!OS) = False And TBAbrir!OS <> "" Then Cmb_OS = TBAbrir!OS
        End If
    End If
    
    If IsNull(TBAbrir!ID_PC) = False And TBAbrir!ID_PC <> "" Then ProcCarregaPC TBAbrir!ID_PC
    
    ProcCarregaListaEmpenhos
    TBAbrir.Close
    
    'Centro de custo
    ProcLimpaCamposCusto False, False
    Frame8.Enabled = False
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto WHERE desenho = '" & txtN_Estoque.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        ProcBloqueiaCampos
        With cmbun
            If TBProduto!Estoque = True Then
                .Locked = True
                .TabStop = False
            Else
                .Locked = False
                .TabStop = True
            End If
        End With
    Else
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos
    End If
End If
CodigoLista1 = Lista.SelectedItem.index
Novo_solicitacao1 = False
Framelista.Enabled = True

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado o código de referência deste produto/serviço."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaPC(ID_PC As Long)
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from tbl_familia where int_codfamilia = " & ID_PC, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Txt_ID_PC = ID_PC
    Txt_conta_contabil = TBFIltro!CODIGO & " - " & TBFIltro!Txt_descricao
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_pedidos, ColumnHeader

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
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select DtValidacao, Status, Data_Autorizacao from Compras_requisicao where ID_Requisicao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    If Cmb_opcao_lista = "Aprovação" Then
                        If IsNull(TBLISTA!DtValidacao) = True Then GoTo Proximo
                        If TBLISTA!status = "CANCELADA" Then GoTo Proximo
                        
                    ElseIf Cmb_opcao_lista = "Validação" Then
                            If IsNull(TBLISTA!Data_autorizacao) = False Then GoTo Proximo
                        Else
                            If Cmb_opcao_lista = "Excluir" And FunVerificaRegistroValidadoSemMsg("Compras_requisicao", "ID_Requisicao = " & ListItems.Item(InitFor), True) = False Then GoTo Proximo
                            If TBLISTA!status <> "ABERTA" And TBLISTA!status <> "CANCELADA" Then GoTo Proximo
                    End If
    
                    If IsNull(TBLISTA!Data_autorizacao) = False Then
                        Set TBAcessos = CreateObject("adodb.recordset")
                        TBAcessos.Open "Select ID_Requisicao from Compras_pedido_lista where ID_Requisicao = " & .ListItems.Item(InitFor) & " and Status_Item <> 'REQUISIT.' and Status_Item <> 'CANCELADO' and Status_Item <> 'NÃO APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAcessos.EOF = False Then
                            GoTo Proximo
                        End If
                        TBAcessos.Close
                    End If
                End If
                TBLISTA.Close
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

Select Case Cmb_opcao_lista
    Case "Excluir": TextoLista = "excluir"
    Case "Status": TextoLista = "alterar status"
    Case "Aprovação": TextoLista = "cancelar aprovação"
    Case "Validação": TextoLista = "cancelar validação"
End Select

With Lista_req
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select DtValidacao, Status, Data_Autorizacao from Compras_requisicao where ID_Requisicao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                If Cmb_opcao_lista = "Aprovação" Then
                    If IsNull(TBLISTA!DtValidacao) = True Then
                        USMsgBox ("Não é permitido aprovar solicitação, pois a mesma ainda não foi validada."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If TBLISTA!status = "CANCELADA" Then
                        USMsgBox ("Não é permitido autorizar/cancelar aprovação, pois o status da solicitação está cancelada."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                ElseIf Cmb_opcao_lista = "Validação" Then
                        If IsNull(TBLISTA!Data_autorizacao) = False Then
                            USMsgBox ("Não é permitido cancelar validação, pois a solicitação já foi aprovada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    Else
                        If IsNull(TBLISTA!DtValidacao) = False And Cmb_opcao_lista = "Excluir" Then
                            USMsgBox ("Não é permitido excluir solicitação, pois a mesma já foi validada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                        If TBLISTA!status <> "ABERTA" And TBLISTA!status <> "CANCELADA" Then
                            USMsgBox ("Não é permitido " & TextoLista & ", pois o status da solicitação está " & TBLISTA!status & "."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                End If
                
                If IsNull(TBLISTA!Data_autorizacao) = False Then
                    Set TBAcessos = CreateObject("adodb.recordset")
                    TBAcessos.Open "Select ID_Requisicao from Compras_pedido_lista where ID_Requisicao = " & .ListItems.Item(InitFor) & " and Status_Item <> 'REQUISIT.' and Status_Item <> 'CANCELADO' and Status_Item <> 'NÃO APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAcessos.EOF = False Then
                        USMsgBox ("Não é permitido " & TextoLista & ", pois os produtos/serviços já sofreram alguma alteração."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBAcessos.Close
                        Exit Sub
                    End If
                    TBAcessos.Close
                End If
            End If
            TBLISTA.Close
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
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_requisicao where id_requisicao = " & Lista_req.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    ProcLimpaCampos
    ProcAbrir
    CodigoLista = Lista_req.SelectedItem.index
    Novo_solicitacao = False
End If
TBCompras.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptQS_com_Click()
On Error GoTo tratar_erro

If OptQS_com.Value = True Then
    With txtQS_com
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
Else
    With txtQS_est
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptQS_est_Click()
On Error GoTo tratar_erro

If OptQS_est.Value = True Then
    With txtQS_est
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
Else
    With txtQS_com
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtNumero = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        Cmb_empresa.Visible = True
        PBLista.Visible = True
        If Lista_req.Visible = True Then Lista_req.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        If SSTab2.Tab = 0 Then PBLista.Visible = False Else PBLista.Visible = True
        If FunVerifNFProdServSemCad(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = False Then Procliberacampos Else ProcBloqueiaCampos
        If Novo_solicitacao = True Then
            USMsgBox ("Salve a solicitação antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            Exit Sub
        End If
        If SSTab2.Tab = 0 Then ProcCarregaLista
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If TXTIDLista = 0 Then
    SSTab2.Tab = 0
    Exit Sub
End If
If SSTab2.Tab = 0 Then
    PBLista.Visible = False
    ProcHabDesabBotoesProdServ
Else
    PBLista.Visible = True
    With USToolBar2
        .ButtonState(3) = 0
        .ButtonState(5) = 5
        .ButtonState(6) = 5
        If SSTab2.Tab = 1 Then
            .ButtonState(2) = 0
            .ButtonState(8) = 0
        Else
            .ButtonState(2) = 5
            .ButtonState(8) = 5
        End If
        .Refresh
    End With
End If

Select Case SSTab2.Tab
    Case 0:
        Lista.SetFocus
        ProcCarregaLista
    Case 1:
        Lista_custo.SetFocus
        If ProcVerifProsseguir = False Then Exit Sub
        ProcCarregaLista_Custo
    Case 2:
'        Lista_empenhos.SetFocus
        If ProcVerifProsseguir = False Then Exit Sub
        Txt_qtde_total_solicitada = txtQS_com

        'Verifica se é requisição de serviço de terceiro
        If Lista.SelectedItem.ListSubItems(13) <> "" Then
            USMsgBox ("Não é permitido fazer o empenho, pois este produto/serviço já está empenhado para uma ordem de produção."), vbExclamation, "CAPRIND v5.0"
            SSTab2.Tab = 0
            Exit Sub
        End If
        ProcCarregaListaEmpenhos
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifProsseguir() As Boolean
On Error GoTo tratar_erro

ProcVerifProsseguir = True
If Novo_solicitacao1 = True Then
    USMsgBox ("Salve o produto/serviço antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab2.Tab = 0
    ProcVerifProsseguir = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub txtN_Estoque_Change()
On Error GoTo tratar_erro

If chkAuto.Value = 0 And chkManual.Value = 0 Then
    txtdescricao = ""
    Txt_descricao_comercial = ""
    cmbun.ListIndex = -1
    Cmb_un_com.ListIndex = -1
    cmbfamilia.ListIndex = -1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarNovoNumero()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Requisicaotexto from Compras_requisicao where Year (Data_Solicitacao) = '" & Year(Date) & "' order by ID_Requisicao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Numero = Left(TBAbrir!Requisicaotexto, Len(TBAbrir!Requisicaotexto) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBAbrir.Close

a = Numero
Ano = Right(Year(Date), 2)
Select Case Len(a)
    Case 1: a = "SOL-0000" & Numero & "/" & Ano
    Case 2: a = "SOL-000" & Numero & "/" & Ano
    Case 3: a = "SOL-00" & Numero & "/" & Ano
    Case 4: a = "SOL-0" & Numero & "/" & Ano
    Case 5: a = "SOL-" & Numero & "/" & Ano
End Select

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

Sub ProcCarregaLista_Req(Pagina As Integer)
On Error GoTo tratar_erro

Lista_req.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
If StrSql_solicitacao = "" Then Exit Sub
Set TBLISTA_Solicitacao = CreateObject("adodb.recordset")
TBLISTA_Solicitacao.Open StrSql_solicitacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Solicitacao.EOF = False Then ProcExibePagina (Pagina)
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista_req.ListItems.Clear
TBLISTA_Solicitacao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Solicitacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Solicitacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Solicitacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Solicitacao.PageSize * (Pagina - 1)), 0), TBLISTA_Solicitacao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Solicitacao.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_req.ListItems
        .Add , , TBLISTA_Solicitacao!ID_Requisicao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Empresa from Empresa where Codigo = " & IIf(IsNull(TBLISTA_Solicitacao!ID_empresa), 0, TBLISTA_Solicitacao!ID_empresa), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Empresa), "", TBAbrir!Empresa)
        End If
        TBAbrir.Close
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Solicitacao!Requisicaotexto), "", TBLISTA_Solicitacao!Requisicaotexto)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Solicitacao!Data_Solicitacao), "", Format(TBLISTA_Solicitacao!Data_Solicitacao, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Solicitacao!solicitado), "", TBLISTA_Solicitacao!solicitado)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Solicitacao!setorsolic), "", TBLISTA_Solicitacao!setorsolic)
        If TBLISTA_Solicitacao!Data_autorizacao <> "" Then
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Solicitacao!Data_autorizacao), "", Format(TBLISTA_Solicitacao!Data_autorizacao, "dd/mm/yy"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Solicitacao!Autorizado), "", TBLISTA_Solicitacao!Autorizado)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Solicitacao!setorautor), "", TBLISTA_Solicitacao!setorautor)
        End If
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Solicitacao!status), "", TBLISTA_Solicitacao!status)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Solicitacao!DtValidacao), "Não", "Sim")
    End With
    TBLISTA_Solicitacao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Solicitacao.RecordCount
If TBLISTA_Solicitacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Solicitacao.PageCount
ElseIf TBLISTA_Solicitacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Solicitacao.PageCount & " de: " & TBLISTA_Solicitacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Solicitacao.AbsolutePage - 1 & " de: " & TBLISTA_Solicitacao.PageCount
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

Private Sub txtOrdem_Change()
On Error GoTo tratar_erro

Cmb_OS.Clear
If txtOrdem <> "" Then
    VerifNumero = txtOrdem
    ProcVerificaNumero
    If VerifNumero = False Then
        txtOrdem = ""
        txtOrdem.SetFocus
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

Private Sub txtOrdem_LostFocus()
On Error GoTo tratar_erro

With txtOrdem
    If .Text <> "" And .Text <> "0" Then
        If FunVerifOPCarregaOS(Cmb_OS, .Text, Novo_solicitacao1, True) = False Then
            .Text = ""
            If Framelista.Enabled = True Then .SetFocus
        End If
    End If
End With
ProcCarregaListaPedidos
    
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
valor = Txt_qtde_total_solicitada
Txt_qtde_total_emp = Format(Valor3, "###,##0.0000")
Txt_qtde_total_disp = Format(valor - Valor3, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaPedidos()
On Error GoTo tratar_erro

Lista_pedidos.ListItems.Clear
If txtOrdem = "" Or txtOrdem = "0" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VC.*, VP.Ncotacao, VP.Revisao, VP.Cliente, PP.ID FROM (vendas_proposta VP INNER JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao) INNER JOIN Producao_pedidos PP on VC.Codigo = PP.IDCarteira where PP.Ordem = " & txtOrdem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_pedidos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!CODIGO
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ncotacao), "", TBLISTA!Ncotacao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Rev_codinterno), "", TBLISTA!Rev_codinterno)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!descricao_tecnica), "", Trim(TBLISTA!descricao_tecnica))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentualCentro_Change()
On Error GoTo tratar_erro

If txtPercentualCentro <> "" Then
    VerifNumero = txtPercentualCentro
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPercentualCentro = ""
        txtPercentualCentro.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentualCentro_LostFocus()
On Error GoTo tratar_erro

txtPercentualCentro = Format(txtPercentualCentro, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtprazo_LostFocus()
On Error GoTo tratar_erro

If txtprazo.Text <> "__/__/____" Then
    VerifData = txtprazo.Text
    ProcVerificaData
    If VerifData = False Then
        txtprazo.Text = "__/__/____"
        txtprazo.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_com_Change()
On Error GoTo tratar_erro

If txtQS_com <> "" Then
    VerifNumero = txtQS_com
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQS_com = ""
        txtQS_com.SetFocus
        Exit Sub
    End If
    If OptQS_com.Value = True Then
        If cmbun <> Cmb_un_com Then
            txtQS_est = FunFormataCasasDecimais(4, FunConversaoFinalUn(cmbun, Cmb_un_com, txtQS_com, txtN_Estoque, True))
        Else
            txtQS_est = FunFormataCasasDecimais(4, txtQS_com)
        End If
    End If
    If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
        txtQS_com_PC = FunCalculaQtdePC(txtN_Estoque, txtQS_com, True, Cmb_un_com)
    Else
        txtQS_com_PC = ""
    End If
Else
    If OptQS_com.Value = True Then txtQS_est = ""
    txtQS_com_PC = ""
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_com_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQS_com

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_com_LostFocus()
On Error GoTo tratar_erro

If txtQS_com <> "" Then txtQS_com = FunFormataCasasDecimais(4, txtQS_com)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoProdutoManual()
On Error GoTo tratar_erro

If cmbun <> "SE" And cmbun <> "SV" And cmbun <> "HS" Then
    txtN_Estoque = FunCriaNovoProdServ(True, "", txtN_Estoque, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 0, "P", "", 0, 0, 0, "", 0, "", "")
Else
    txtN_Estoque = FunCriaNovoProdServ(True, "", txtN_Estoque, txtreferencia, 0, txtdescricao, Txt_descricao_comercial, cmbfamilia, 0, 0, 0, cmbun, Cmb_un_com, 0, True, False, True, False, 5, "S", "", 0, 0, 0, "", 0, "", "")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_est_Change()
On Error GoTo tratar_erro

If txtQS_est <> "" Then
    If OptQS_est.Value = True Then
        VerifNumero = txtQS_est
        ProcVerificaNumero
        If VerifNumero = False Then
            txtQS_est = ""
            txtQS_est.SetFocus
            Exit Sub
        End If
        If cmbun <> Cmb_un_com Then
            txtQS_com = FunFormataCasasDecimais(4, FunConversaoFinalUn(cmbun, Cmb_un_com, txtQS_est, txtN_Estoque, False))
        Else
            txtQS_com = FunFormataCasasDecimais(4, txtQS_est)
        End If
    End If
Else
    If OptQS_est.Value = True Then txtQS_com = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_est_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQS_est

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQS_est_LostFocus()
On Error GoTo tratar_erro

If txtQS_est <> "" Then txtQS_est = FunFormataCasasDecimais(4, txtQS_est)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtstatus_Change()
On Error GoTo tratar_erro

If txtStatus.Text = "CANCELADA" Then Cmd_dados_cancelamento.Enabled = True Else Cmd_dados_cancelamento.Enabled = False
    
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
    Case 8: ProcStatus
    Case 9: ProcCopiar
    Case 10: ProcValidarRegistros Lista_req, "Outros/Solicitação"
    Case 11: ProcValidarRegistros Lista_req, "Outros/Solicitação/Autorizar solicitação"
    Case 12: ProcAtualizar
    Case 14: ProcAjuda
    Case 15: ProcSair
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
    Case 7: ProcStatus1
    Case 8: ProcCopiar_CC
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifUsuario(MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifUsuario = True
If txtSolicitado <> pubUsuario Then
    If MostrarMsg = True Then USMsgBox ("Só é permitido modificação na solicitação pelo usuário " & txtSolicitado.Text & "."), vbExclamation, "CAPRIND v5.0"
    ProcVerifUsuario = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Function ProcVerifSatus(Acao As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifSatus = True
If txtStatus <> "ABERTA" Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & ", pois a solicitação está " & txtStatus & "."), vbExclamation, "CAPRIND v5.0"
    ProcVerifSatus = False
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
