VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInstrumentos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Instrumentos"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmInstrumentos.frx":0000
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
      TabIndex        =   72
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
   Begin VB.ComboBox cmbStatus 
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
      ItemData        =   "frmInstrumentos.frx":000C
      Left            =   12960
      List            =   "frmInstrumentos.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Status."
      Top             =   1695
      Width           =   2130
   End
   Begin VB.ComboBox cmbTipo 
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
      ItemData        =   "frmInstrumentos.frx":006F
      Left            =   9120
      List            =   "frmInstrumentos.frx":0085
      Style           =   2  'Dropdown List
      TabIndex        =   22
      ToolTipText     =   "Tipo da resolução do equipamento"
      Top             =   4090
      Width           =   1920
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Dados do instrumento"
      TabPicture(0)   =   "frmInstrumentos.frx":00D6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Framegerais"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtid"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Txt_ID_estoque"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Dados da calibração"
      TabPicture(1)   =   "frmInstrumentos.frx":00F2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(1)=   "Lista1"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "txtid_afericao"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Histórico de utilização"
      TabPicture(2)   =   "frmInstrumentos.frx":010E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(1)=   "Lista2"
      Tab(2).ControlCount=   2
      Begin VB.TextBox Txt_ID_estoque 
         Height          =   315
         Left            =   1620
         TabIndex        =   78
         Top             =   5160
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   68
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
            Left            =   3750
            TabIndex        =   23
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   28
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmInstrumentos.frx":012A
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
            TabIndex        =   27
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmInstrumentos.frx":38CE
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
            Left            =   10680
            TabIndex        =   26
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmInstrumentos.frx":73D7
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
            TabIndex        =   29
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmInstrumentos.frx":B4C6
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
            Left            =   4380
            TabIndex        =   80
            Top             =   240
            Width           =   1440
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
            TabIndex        =   71
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
            TabIndex        =   70
            Top             =   240
            Width           =   1095
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
            Left            =   3060
            TabIndex        =   69
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox txtid 
         Height          =   315
         Left            =   1200
         TabIndex        =   61
         Text            =   "0"
         Top             =   5160
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtid_afericao 
         Height          =   315
         Left            =   -73080
         TabIndex        =   60
         Text            =   "0"
         Top             =   4980
         Visible         =   0   'False
         Width           =   405
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
         Height          =   2625
         Left            =   -74925
         TabIndex        =   54
         Top             =   1305
         Width           =   15195
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frmInstrumentos.frx":ED52
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Limpar caminho."
            Top             =   2190
            Width           =   315
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmInstrumentos.frx":EE90
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Visualizar arquivo."
            Top             =   2190
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
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Caminho da imagem do certificado."
            Top             =   2190
            Width           =   13845
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14040
            Picture         =   "frmInstrumentos.frx":F452
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Localizar comprovante."
            Top             =   2190
            Width           =   315
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
            Left            =   1590
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   13425
         End
         Begin VB.TextBox txtData1 
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
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1395
         End
         Begin VB.Frame FrameAprovado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprovado*"
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
            Height          =   1005
            Left            =   13950
            TabIndex        =   55
            Top             =   900
            Width           =   1065
            Begin VB.OptionButton optNao 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Não"
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
               Height          =   210
               Left            =   150
               TabIndex        =   37
               Top             =   660
               Width           =   585
            End
            Begin VB.OptionButton optSim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Sim"
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
               Height          =   210
               Left            =   150
               TabIndex        =   36
               Top             =   345
               Width           =   555
            End
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
            Left            =   1590
            MaxLength       =   255
            TabIndex        =   35
            ToolTipText     =   "Certificado."
            Top             =   1590
            Width           =   12240
         End
         Begin VB.TextBox txtOrgao 
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
            MaxLength       =   255
            TabIndex        =   33
            ToolTipText     =   "Orgão."
            Top             =   990
            Width           =   12240
         End
         Begin MSComCtl2.DTPicker txtAferido 
            Height          =   315
            Left            =   180
            TabIndex        =   32
            ToolTipText     =   "Data de calibração."
            Top             =   990
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
            Format          =   171638785
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtProxima_Afericao 
            Height          =   315
            Left            =   180
            TabIndex        =   34
            ToolTipText     =   "Data da próxima calibração."
            Top             =   1590
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
            Format          =   171638785
            CurrentDate     =   39057
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho da imagem do certificado"
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
            Left            =   5880
            TabIndex        =   67
            Top             =   1980
            Width           =   2445
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
            Index           =   8
            Left            =   7845
            TabIndex        =   64
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
            Index           =   6
            Left            =   705
            TabIndex        =   63
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orgão*"
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
            Left            =   7440
            TabIndex        =   59
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Certificado*"
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
            Left            =   7275
            TabIndex        =   58
            Top             =   1380
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Próx. calibração"
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
            Left            =   300
            TabIndex        =   57
            Top             =   1380
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data calibração"
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
            Left            =   322
            TabIndex        =   56
            Top             =   780
            Width           =   1110
         End
      End
      Begin VB.Frame Framegerais 
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
         Height          =   3255
         Left            =   75
         TabIndex        =   45
         Top             =   1305
         Width           =   15195
         Begin VB.ComboBox cmbFREQ 
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
            ItemData        =   "frmInstrumentos.frx":F554
            Left            =   7860
            List            =   "frmInstrumentos.frx":F56D
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Frequência"
            Top             =   2790
            Width           =   1170
         End
         Begin VB.TextBox txtEIM 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            ToolTipText     =   "ERRO TOTAL (Erro máximo de indicação + Incerteza de Medição)"
            Top             =   2790
            Width           =   885
         End
         Begin VB.TextBox txtINM 
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
            Left            =   6060
            TabIndex        =   19
            ToolTipText     =   $"frmInstrumentos.frx":F5AB
            Top             =   2790
            Width           =   885
         End
         Begin VB.TextBox txtEMI 
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
            Left            =   5160
            TabIndex        =   18
            ToolTipText     =   $"frmInstrumentos.frx":F5D7
            Top             =   2790
            Width           =   885
         End
         Begin VB.TextBox txtERM 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            ToolTipText     =   "ERRO MÁXIMO (1/3IT)"
            Top             =   2790
            Width           =   885
         End
         Begin VB.TextBox txtINT 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            ToolTipText     =   "IT (Intervalo de Tolerância)"
            Top             =   2790
            Width           =   1095
         End
         Begin VB.TextBox txtDPA 
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
            Left            =   1980
            TabIndex        =   17
            ToolTipText     =   "Desvio padrão"
            Top             =   2790
            Width           =   1155
         End
         Begin VB.TextBox txtRes 
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
            Left            =   1080
            TabIndex        =   16
            ToolTipText     =   "Resolução do equipamento"
            Top             =   2790
            Width           =   885
         End
         Begin VB.TextBox txtFMT 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
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
            Left            =   210
            TabIndex        =   15
            ToolTipText     =   "Faixa de medição"
            Top             =   2790
            Width           =   885
         End
         Begin VB.TextBox Txt_cod_ref 
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
            Left            =   2180
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   390
            Width           =   2120
         End
         Begin VB.TextBox Txt_numero_serie 
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
            Left            =   4305
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Número de série."
            Top             =   390
            Width           =   2520
         End
         Begin VB.TextBox txtDescricao 
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
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   7965
         End
         Begin VB.CommandButton Cmd_limpar_LA 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8490
            Picture         =   "frmInstrumentos.frx":F607
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Limpar local de armazenamento."
            Top             =   1590
            Width           =   315
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
            Left            =   8100
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   4765
         End
         Begin VB.CommandButton cmdnome 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8160
            Picture         =   "frmInstrumentos.frx":F745
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Localizar local de armazenamento."
            Top             =   1590
            Width           =   315
         End
         Begin VB.TextBox txtLocal_armaz 
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
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Local de armazenamento."
            Top             =   1590
            Width           =   7965
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1245
         End
         Begin VB.TextBox cmbFamilia 
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
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   990
            Width           =   6855
         End
         Begin VB.TextBox txtParametro 
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
            Height          =   915
            Left            =   11010
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            ToolTipText     =   "Observações para calibração."
            Top             =   2190
            Width           =   4035
         End
         Begin VB.TextBox txtFabricante 
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
            Left            =   1590
            MaxLength       =   255
            TabIndex        =   13
            ToolTipText     =   "Fabricante."
            Top             =   2190
            Width           =   9360
         End
         Begin VB.TextBox txtNumero 
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1980
         End
         Begin MSComCtl2.DTPicker txtData_Aquisicao 
            Height          =   315
            Left            =   180
            TabIndex        =   12
            ToolTipText     =   "Data de aquisição."
            Top             =   2190
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
            Format          =   171704321
            CurrentDate     =   39057
         End
         Begin VB.TextBox txtFuncionario 
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
            Left            =   8910
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Portador do instrumento."
            Top             =   1590
            Width           =   6105
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Frequência"
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
            Left            =   7995
            TabIndex        =   92
            Top             =   2580
            Width           =   795
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Erro Total"
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
            Left            =   7050
            TabIndex        =   91
            Top             =   2580
            Width           =   705
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   6412
            TabIndex        =   89
            Top             =   2580
            Width           =   180
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "EMI"
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
            Left            =   5460
            TabIndex        =   88
            Top             =   2580
            Width           =   270
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Erro Max."
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
            Left            =   4350
            TabIndex        =   87
            Top             =   2580
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Int.tolerância"
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
            Left            =   3240
            TabIndex        =   85
            Top             =   2580
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Desvio Padrão"
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
            TabIndex        =   83
            Top             =   2580
            Width           =   1035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Resolução"
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
            Left            =   1170
            TabIndex        =   82
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Faixa Med."
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
            Left            =   270
            TabIndex        =   81
            Top             =   2580
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
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
            TabIndex        =   79
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Portador do instrumento"
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
            Left            =   11085
            TabIndex        =   77
            Top             =   1380
            Width           =   1755
         End
         Begin VB.Label Label3 
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
            Index           =   15
            Left            =   3817
            TabIndex        =   76
            Top             =   780
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo resolução"
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
            Left            =   9420
            TabIndex        =   66
            Top             =   2580
            Width           =   1035
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
            Index           =   9
            Left            =   13665
            TabIndex        =   65
            Top             =   180
            Width           =   465
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
            Index           =   1
            Left            =   10020
            TabIndex        =   62
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Número de série"
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
            Left            =   4980
            TabIndex        =   53
            Top             =   180
            Width           =   1170
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   3270
            TabIndex        =   52
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações de calibração"
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
            Left            =   12345
            TabIndex        =   51
            Top             =   1980
            Width           =   1935
         End
         Begin VB.Label Label2 
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
            Index           =   5
            Left            =   11347
            TabIndex        =   50
            Top             =   780
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
            Index           =   4
            Left            =   7290
            TabIndex        =   49
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante*"
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
            Left            =   6485
            TabIndex        =   48
            Top             =   1980
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. aquisição"
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
            Left            =   405
            TabIndex        =   47
            Top             =   1980
            Width           =   930
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   555
            TabIndex        =   46
            Top             =   180
            Width           =   1230
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   4530
         Left            =   75
         TabIndex        =   21
         Top             =   4545
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   7990
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
            Object.Width           =   512
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. de ref."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "N. de série"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   6006
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Dt. aquisição"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Fabricante"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView Lista1 
         Height          =   5775
         Left            =   -74925
         TabIndex        =   42
         Top             =   3945
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10186
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
            Text            =   "Data calibração"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Orgão"
            Object.Width           =   8295
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Proxima calibração"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Certificado"
            Object.Width           =   8295
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Aprovado"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lista2 
         Height          =   8400
         Left            =   -74925
         TabIndex        =   43
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   14817
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "T"
            Text            =   "Plano de mediçao"
            Object.Width           =   3175
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
            Object.Width           =   12515
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "N° de rastreabilidade"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Inspetor"
            Object.Width           =   4939
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   73
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
         ButtonToolTipText8=   "Status (F8)"
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
         ButtonCaption9  =   "Atualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Utilizado pelo administrador do sistema."
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
         ButtonWidth9    =   59
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
         ButtonLeft10    =   461
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   465
         ButtonTop11     =   2
         ButtonWidth11   =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   508
         ButtonTop12     =   2
         ButtonWidth12   =   30
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
         ButtonLeft13    =   540
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   13050
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmInstrumentos.frx":F847
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   74
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
         ButtonCaption7  =   "Atualizar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Utilizado pelo administrador do sistema."
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
         ButtonWidth7    =   59
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
         ButtonLeft8     =   370
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   374
         ButtonTop9      =   2
         ButtonWidth9    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft10    =   417
         ButtonTop10     =   2
         ButtonWidth10   =   30
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   5
         ButtonLeft11    =   449
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   12330
            Top             =   150
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13050
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmInstrumentos.frx":16D0D
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   75
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
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
         ButtonCaption1  =   "Relatório"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Relatório (F5)"
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
         ButtonWidth1    =   60
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Anterior"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Registro anterior."
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
         ButtonLeft2     =   64
         ButtonTop2      =   2
         ButtonWidth2    =   55
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Próximo"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Próximo registro."
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
         ButtonLeft3     =   121
         ButtonTop3      =   2
         ButtonWidth3    =   55
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
         ButtonLeft4     =   178
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
         ButtonLeft5     =   182
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
         ButtonLeft6     =   225
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
         ButtonLeft7     =   257
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         ButtonUseMaskColor7=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   13050
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmInstrumentos.frx":1CDF4
            Count           =   1
         End
      End
   End
End
Attribute VB_Name = "frmInstrumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Instrumentos    As Boolean 'OK
Dim Novo_Instrumentos1      As Boolean 'OK
Public StrSql_Instrumentos_Localizar As String 'OK
Public FormulaRel_Instrumentos As String 'OK
Dim TBLISTA_Instrumentos    As ADODB.Recordset 'OK

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro
            
lblRegistros.Caption = "Nº de registros: 0"
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If StrSql_Instrumentos_Localizar = "" Then Exit Sub
Set TBLISTA_Instrumentos = CreateObject("adodb.recordset")
TBLISTA_Instrumentos.Open StrSql_Instrumentos_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Instrumentos.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Instrumentos.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Instrumentos.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Instrumentos.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Instrumentos.RecordCount - IIf(Pagina > 1, (TBLISTA_Instrumentos.PageSize * (Pagina - 1)), 0), TBLISTA_Instrumentos.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Instrumentos.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Instrumentos!CODIGO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Instrumentos!Numero), "", TBLISTA_Instrumentos!Numero)
        If IsNull(TBLISTA_Instrumentos!Ref) = True Or TBLISTA_Instrumentos!Ref = "" Then .Item(.Count).SubItems(2) = FunCarregaCodRef(TBLISTA_Instrumentos!Numero) Else .Item(.Count).SubItems(2) = TBLISTA_Instrumentos!Ref
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Instrumentos!Numero_serie), "", TBLISTA_Instrumentos!Numero_serie)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Instrumentos!Descricao), "", TBLISTA_Instrumentos!Descricao)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Instrumentos!Data_Aquisicao), "", Format(TBLISTA_Instrumentos!Data_Aquisicao, "dd/mm/yy"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Instrumentos!Fabricante), "", TBLISTA_Instrumentos!Fabricante)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Instrumentos!Familia), "", TBLISTA_Instrumentos!Familia)
    End With
    TBLISTA_Instrumentos.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Instrumentos.RecordCount
If TBLISTA_Instrumentos.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Instrumentos.PageCount
ElseIf TBLISTA_Instrumentos.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Instrumentos.PageCount & " de: " & TBLISTA_Instrumentos.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Instrumentos.AbsolutePage - 1 & " de: " & TBLISTA_Instrumentos.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from instrumentos where numero <> 'Null' order by numero", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("numero = '" & txtNumero & "'")
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtNumero = TBLISTA!Numero
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select I.*, EC.Numero_serie, EC.ref from Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.numero = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        ProcLimpaCamposAfericao
        ProcPuxaDados
        ProcCarregaListaAfericao
        ProcCarregaListaHistorico
    Else
        USMsgBox ("Fim dos cadastros de instrumentos."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Instrumentos1 = False

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
Acao = "alterar o status"
If txtNumero.Text = "" Then
    NomeCampo = "o instrumento"
    ProcVerificaAcao
    Exit Sub
End If
If Novo_Instrumentos = True Then
    USMsgBox ("Salve o instrumento antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmInstrumentos_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

txt_Caminho = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_LA_Click()
On Error GoTo tratar_erro

txtLocal_armaz = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "CQ_instrumentos.rpt"
ProcImprimirRel FormulaRel_Instrumentos, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdnome_Click()
On Error GoTo tratar_erro
  
frmInstrumentos_localarmaz.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId.Text = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from instrumentos where numero <> 'Null' order by numero", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("numero = '" & txtNumero & "'")
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtNumero = TBLISTA!Numero
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select I.*, EC.Numero_serie, EC.ref from Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.numero = '" & txtNumero & "'", Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpar
        ProcLimpaCamposAfericao
        ProcPuxaDados
        ProcCarregaListaAfericao
        ProcCarregaListaHistorico
    Else
        USMsgBox ("Fim dos cadastros de instrumentos."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_Instrumentos1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Instrumentos.AbsolutePage <> 2 Then
    If TBLISTA_Instrumentos.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Instrumentos.PageCount - 1)
    Else
        TBLISTA_Instrumentos.AbsolutePage = TBLISTA_Instrumentos.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Instrumentos.AbsolutePage)
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
    TBLISTA_Instrumentos.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Instrumentos.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Instrumentos.AbsolutePage = 1
ProcExibePagina (TBLISTA_Instrumentos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Instrumentos.AbsolutePage <> -3 Then
    If TBLISTA_Instrumentos.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Instrumentos.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Instrumentos.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Instrumentos.AbsolutePage = TBLISTA_Instrumentos.PageCount
ProcExibePagina (TBLISTA_Instrumentos.AbsolutePage)

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
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcStatus
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_Afericao
            Case vbKeyF3: ProcSalvar_Afericao
            Case vbKeyF4: ProcExcluir_Afericao
            Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF5: ProcImprimir
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

ProcCarregaToolBar1 Me, 15200, 13, True
ProcCarregaToolBar2 Me, 15200, 11, True
ProcCarregaToolBar3 Me, 15200, 7, True
Formulario = "Qualidade/Instrumentos"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from afericao where proxima_afericao = '" & Format(Date, "Short Date") & "' order by proxima_afericao", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
   USMsgBox ("Existem instrumentos com data de calibração agendada para hoje."), vbInformation, "CAPRIND v5.0"
   frmAgenda_afericao.Show 1
End If
TBFI.Close

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Instrumentos"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362I" Then
    If USMsgBox("Deseja realmente atualizar os dados no módulo de engenharia?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBInstrumentos = CreateObject("adodb.recordset")
        TBInstrumentos.Open "Select * from instrumentos", Conexao, adOpenKeyset, adLockOptimistic
        If TBInstrumentos.EOF = False Then
            TBInstrumentos.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBInstrumentos.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBInstrumentos.MoveFirst
            Do While TBInstrumentos.EOF = False
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "Select * from projproduto where desenho = '" & TBInstrumentos!Numero & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = True Then TBProduto.AddNew
                TBProduto!Desenho = TBInstrumentos!Numero
                TBProduto!descricaotecnica = IIf(IsNull(TBInstrumentos!Descricao), "", TBInstrumentos!Descricao)
                TBProduto!RevDesenho = 0
                TBProduto!Data = Date
                TBProduto!Descricao = IIf(IsNull(TBInstrumentos!Descricao), "", TBInstrumentos!Descricao)
                TBProduto!Classe = IIf(IsNull(TBInstrumentos!Familia), "", TBInstrumentos!Familia)
                TBProduto!Unidade = "PÇ"
                TBProduto!Responsavel = pubUsuario
                TBProduto!Compras = True
                TBProduto!Producao = True
                TBProduto!Qualidade = True
                TBProduto!nacional = True
                TBProduto.Update
                TBProduto.Close
                TBInstrumentos.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBInstrumentos.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Instrumentos"
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

Private Sub ProcAtualiza_Afericao()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362I1" Then
    If USMsgBox("Deseja realmente atualizar as datas das aferições nos instrumentos?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBInstrumentos = CreateObject("adodb.recordset")
        TBInstrumentos.Open "Select * from instrumentos order by numero", Conexao, adOpenKeyset, adLockOptimistic
        If TBInstrumentos.EOF = False Then
            TBInstrumentos.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBInstrumentos.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBInstrumentos.MoveFirst
            Do While TBInstrumentos.EOF = False
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select * from Afericao where Numero = '" & TBInstrumentos!Numero & "' order by Proxima_afericao", Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    TBAfericao.MoveLast
                    TBInstrumentos!ID_ultima_afericao = TBAfericao!CODIGO
                    TBInstrumentos!Data_ultima_afericao = TBAfericao!Aferido
                    TBInstrumentos!Data_proxima_afericao = TBAfericao!Proxima_afericao
                Else
                    TBInstrumentos!ID_ultima_afericao = Null
                    TBInstrumentos!Data_ultima_afericao = Null
                    TBInstrumentos!Data_proxima_afericao = Null
                End If
                TBInstrumentos.Update
                TBAfericao.Close
                TBInstrumentos.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBInstrumentos.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Instrumentos"
        Evento = "Atualizar1"
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
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) instrumento(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Instrumentos where Codigo = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from afericao where ID_inst = " & .ListItems(InitFor)
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Codigo from Instrumentos where numero = '" & .ListItems(InitFor).ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then TextoFiltro = "Instrumento = 'False'" Else TextoFiltro = "Instrumento = 'True'"
            TBAbrir.Close
            Conexao.Execute "UPDATE projproduto Set " & TextoFiltro & " where Desenho = '" & .ListItems(InitFor).ListSubItems(1) & "'"
            
            '==================================
            Modulo = "Qualidade/Instrumentos"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            If .ListItems(InitFor).ListSubItems(2) <> "" Then Documento = Documento & " - Cód. de ref.: " & .ListItems(InitFor).ListSubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) instrumento(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Instrumento(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpar
    ProcCarregaLista (1)
    Framegerais.Enabled = False
    cmbStatus.Enabled = False
    cmbTipo.Enabled = False
    Novo_Instrumentos = False
    ProcLimparTudo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_Afericao()
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
                If USMsgBox("Deseja realmente excluir esta(s) calibração(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Afericao where Codigo = " & .ListItems(InitFor)
            '==================================
            Modulo = "Qualidade/Instrumentos"
            Evento = "Excluir calibração"
            ID_documento = .ListItems(InitFor)
            Documento = "Cód. interno: " & txtNumero
            Documento1 = "Data calibração: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) calibração(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Calibração(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposAfericao
    ProcCarregaListaAfericao
    Frame3.Enabled = False
    ProcVerificaUltimaAfericao
    Novo_Instrumentos1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmInstrumentos_abrir.Show 1

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
ProcLimpar
frmInstrumentos_localizaritem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame3.Enabled = False
ProcLimpaCamposAfericao
Novo_Instrumentos1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Afericao()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ProcLimpaCamposAfericao
Novo_Instrumentos1 = True
Frame3.Enabled = True
cmbStatus.Locked = False
cmbStatus.TabStop = True
txtAferido.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Instrumentos = True Then
    If USMsgBox("O instrumento ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Instrumentos = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Instrumentos1 = True Then
    If USMsgBox("A calibração ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar_Afericao
        If Novo_Instrumentos1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Instrumentos = False
Novo_Instrumentos1 = False
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
If Framegerais.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtFabricante.Text = "" Then
    USMsgBox ("Informe o fabricante antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtFabricante.SetFocus
    Exit Sub
End If
Set TBInstrumentos = CreateObject("adodb.recordset")
TBInstrumentos.Open "Select * from Instrumentos where Codigo = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBInstrumentos.EOF = True Then
    TBInstrumentos.AddNew
    TBInstrumentos!Bloqueado = False
End If
If txtData = "" Then TBInstrumentos!Data = Date Else TBInstrumentos!Data = txtData
If txtResponsavel = "" Then TBInstrumentos!Responsavel = pubUsuario Else TBInstrumentos!Responsavel = txtResponsavel
TBInstrumentos!Numero = txtNumero
TBInstrumentos!Descricao = txtdescricao.Text
TBInstrumentos!Familia = cmbfamilia.Text
TBInstrumentos!Data_Aquisicao = txtData_Aquisicao.Value
TBInstrumentos!Fabricante = txtFabricante.Text
TBInstrumentos!parametros = txtParametro
TBInstrumentos!local_armaz = txtLocal_armaz
TBInstrumentos!status = cmbStatus
TBInstrumentos!Tipo = cmbTipo
TBInstrumentos!IDEstoque = IIf(Txt_ID_estoque = "", 0, Txt_ID_estoque)
TBInstrumentos!FMT = IIf(txtFMT.Text = "", 0, txtFMT.Text)
TBInstrumentos!RES = IIf(txtRes.Text = "", 0, txtRes.Text)
TBInstrumentos!DPA = IIf(txtDPA.Text = "", 0, txtDPA.Text)
TBInstrumentos!Int = IIf(txtINT.Text = "", 0, txtINT.Text)
TBInstrumentos!ERM = IIf(txtERM.Text = "", 0, txtERM.Text)
TBInstrumentos!EMI = IIf(txtEMI.Text = "", 0, txtEMI.Text)
TBInstrumentos!INM = IIf(txtINM.Text = "", 0, txtINM.Text)
TBInstrumentos!EIM = IIf(txtEIM.Text = "", 0, txtEIM.Text)
TBInstrumentos!Tipo = cmbTipo.Text
TBInstrumentos!FRQ = cmbFREQ.Text

TBInstrumentos.Update
txtId = TBInstrumentos!CODIGO
TBInstrumentos.Close
If Novo_Instrumentos = True Then
    USMsgBox ("Novo instrumento cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_Instrumentos_Localizar = "Select I.CODIGO, I.Numero, EC.ref, EC.Numero_serie, I.Descricao, I.Data_Aquisicao, I.Fabricante, I.Familia from Instrumentos I INNER JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.Codigo = " & txtId
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (1)
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Instrumentos"
ID_documento = txtId
Documento = "Cód. interno: " & txtNumero
Documento1 = ""
ProcGravaEvento
'==================================
Conexao.Execute "UPDATE projproduto Set Instrumento = 'True' where Desenho = '" & txtNumero & "'"
Novo_Instrumentos = False
                
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpar()
On Error GoTo tratar_erro

txtId = 0
txtNumero.Text = ""
Txt_cod_ref = ""
Txt_numero_serie = ""
cmbfamilia.Text = ""
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
cmbStatus.ListIndex = -1
txtdescricao.Text = ""
txtLocal_armaz = ""
txtFuncionario = ""
txtData_Aquisicao.Value = Date
txtFabricante.Text = ""
txtParametro = ""
cmbTipo.ListIndex = -1
Txt_ID_estoque = ""
CodigoLista = 0
Caption = "Controle de qualidade - Instrumentos"
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId = TBProduto!CODIGO
txtNumero.Text = IIf(IsNull(TBProduto!Numero), "", TBProduto!Numero)
If IsNull(TBProduto!Ref) = True Or TBProduto!Ref = "" Then Txt_cod_ref = FunCarregaCodRef(TBProduto!Numero) Else Txt_cod_ref = TBProduto!Ref
Txt_numero_serie = IIf(IsNull(TBProduto!Numero_serie), "", TBProduto!Numero_serie)

Caption = "Controle de qualidade - Instrumentos (Cód. interno : " & TBProduto!Numero & ")"
txtData = IIf(IsNull(TBProduto!Data), "", Format(TBProduto!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBProduto!Responsavel), "", TBProduto!Responsavel)

With cmbStatus
    If TBProduto!Bloqueado = True Then
        .Text = "Bloqueado"
        .Locked = True
        .TabStop = False
    Else
        .Locked = False
        .TabStop = True
        If IsNull(TBProduto!status) = False And TBProduto!status <> "" Then .Text = TBProduto!status
    End If
End With
txtdescricao.Text = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
txtData_Aquisicao.Value = IIf(IsNull(TBProduto!Data_Aquisicao), Date, Format(TBProduto!Data_Aquisicao, "dd/mm/yy"))
txtFabricante.Text = IIf(IsNull(TBProduto!Fabricante), "", TBProduto!Fabricante)
cmbfamilia = IIf(IsNull(TBProduto!Familia), "", TBProduto!Familia)
txtParametro = IIf(IsNull(TBProduto!parametros), "", TBProduto!parametros)
txtLocal_armaz = IIf(IsNull(TBProduto!local_armaz), "", TBProduto!local_armaz)
If IsNull(TBProduto!Tipo) = False And TBProduto!Tipo <> "" Then cmbTipo = TBProduto!Tipo
Txt_ID_estoque = IIf(IsNull(TBProduto!IDEstoque), "", TBProduto!IDEstoque)

'============================================================================
txtFMT.Text = IIf(IsNull(TBProduto!FMT), "", TBProduto!FMT)
txtRes.Text = IIf(IsNull(TBProduto!RES), "", TBProduto!RES)
txtDPA.Text = IIf(IsNull(TBProduto!DPA), "", TBProduto!DPA)
txtINT.Text = IIf(IsNull(TBProduto!Int), "", TBProduto!Int)
txtERM.Text = IIf(IsNull(TBProduto!ERM), "", TBProduto!ERM)
txtEMI.Text = IIf(IsNull(TBProduto!EMI), "", TBProduto!EMI)
txtINM.Text = IIf(IsNull(TBProduto!INM), "", TBProduto!INM)
txtEIM.Text = IIf(IsNull(TBProduto!EIM), "", TBProduto!EIM)

If TBProduto!FRQ <> "" And TBProduto!FRQ <> 0 Then
cmbFREQ.Text = IIf(IsNull(TBProduto!FRQ), "", TBProduto!FRQ)
End If

'============================================================================

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Funcionario from CFI C INNER JOIN Estoque_Controle E ON C.idestoque = E.idestoque where C.Codigo_Produto = '" & txtNumero & "' AND E.Numero_serie = '" & Txt_numero_serie & "' AND C.Status = 'EM ABERTO'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then txtFuncionario = TBAbrir!Funcionario
TBAbrir.Close

Framegerais.Enabled = True
cmbStatus.Enabled = True
cmbTipo.Enabled = True
Novo_Instrumentos = False
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar_Afericao()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtOrgao = "" Then
    NomeCampo = "o orgão"
    ProcVerificaAcao
    txtOrgao.SetFocus
    Exit Sub
End If
If txtCertificado = "" Then
    NomeCampo = "o certificado"
    ProcVerificaAcao
    txtCertificado.SetFocus
    Exit Sub
End If
If optSim.Value = False And optNao.Value = False Then
    NomeCampo = "se a calibração foi aprovada ou não"
    ProcVerificaAcao
    Exit Sub
End If
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Afericao where Codigo = " & txtid_afericao, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = True Then TBAfericao.AddNew
If txtData1 = "" Then TBAfericao!Data = Date Else TBAfericao!Data = txtData1
If txtResponsavel1 = "" Then TBAfericao!Responsavel = pubUsuario Else TBAfericao!Responsavel = txtResponsavel1
TBAfericao!ID_inst = txtId
TBAfericao!Data_Aquisicao = txtData_Aquisicao.Value
TBAfericao!Fabricante = txtFabricante.Text
TBAfericao!Aferido = txtAferido.Value
TBAfericao!Orgao = txtOrgao.Text
TBAfericao!Proxima_afericao = txtProxima_Afericao.Value
TBAfericao!Certificado = txtCertificado.Text
TBAfericao!caminho = txt_Caminho
If optSim.Value = True Then TBAfericao!Aprovado = True Else TBAfericao!Aprovado = False
TBAfericao.Update
txtid_afericao = TBAfericao!CODIGO
TBAfericao.Close
ProcCarregaListaAfericao
If Novo_Instrumentos1 = True Then
    USMsgBox ("Nova calibração cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova calibração"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar calibração"
    If CodigoLista1 <> 0 And Lista1.ListItems.Count <> 0 Then
        Lista1.SelectedItem = Lista1.ListItems(CodigoLista1)
        Lista1.SetFocus
    End If
End If
'==================================
Modulo = "Qualidade/Instrumentos"
ID_documento = txtid_afericao
Documento = "Cód. interno: " & txtNumero
Documento1 = "Data calibração: " & txtAferido
ProcGravaEvento
'==================================
Novo_Instrumentos1 = False
ProcVerificaUltimaAfericao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaUltimaAfericao()
On Error GoTo tratar_erro

Set TBInstrumentos = CreateObject("adodb.recordset")
TBInstrumentos.Open "Select ID_ultima_afericao, Data_ultima_afericao, Data_proxima_afericao from Instrumentos where Codigo = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBInstrumentos.EOF = False Then
    Set TBAfericao = CreateObject("adodb.recordset")
    TBAfericao.Open "Select CODIGO, Aferido, Proxima_afericao from Afericao where ID_inst = " & txtId & " order by Proxima_afericao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAfericao.EOF = False Then
        TBAfericao.MoveLast
        TBInstrumentos!ID_ultima_afericao = TBAfericao!CODIGO
        TBInstrumentos!Data_ultima_afericao = TBAfericao!Aferido
        TBInstrumentos!Data_proxima_afericao = TBAfericao!Proxima_afericao
    Else
        TBInstrumentos!ID_ultima_afericao = Null
        TBInstrumentos!Data_ultima_afericao = Null
        TBInstrumentos!Data_proxima_afericao = Null
    End If
    TBInstrumentos.Update
    TBAfericao.Close
End If
TBInstrumentos.Close

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
                ProcVerificaRegistroUtilizadoSemMsg "Medicaodimensao_instrumentos", "ID_inst = " & .ListItems(InitFor)
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
            Mensagem = "Não é permitido excluir este instrumento, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "Medicaodimensao_instrumentos", "ID_inst = " & .ListItems(InitFor), "Qualidade/Controle de medição"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
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
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select I.*, EC.Numero_serie, EC.ref from Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcLimpar
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBProduto.Close

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
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
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

Private Sub Lista1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista1.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Afericao where Codigo = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposAfericao
    txtid_afericao = TBLISTA!CODIGO
    txtData1 = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
    txtResponsavel1 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    If IsNull(TBLISTA!Aferido) = False Then txtAferido.Value = TBLISTA!Aferido
    txtOrgao = IIf(IsNull(TBLISTA!Orgao), "", TBLISTA!Orgao)
    If IsNull(TBLISTA!Proxima_afericao) = False Then txtProxima_Afericao.Value = TBLISTA!Proxima_afericao
    txtCertificado = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
    If TBLISTA!Aprovado = True Then optSim.Value = True Else optNao.Value = True
    txt_Caminho = IIf(IsNull(TBLISTA!caminho), "", TBLISTA!caminho)
End If
TBLISTA.Close
Frame3.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista2, ColumnHeader

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
        cmbStatus.Visible = True
        cmbTipo.Visible = True
        Lista.SetFocus
    Case 1:
        cmbStatus.Visible = False
        cmbTipo.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        ProcCarregaListaAfericao
    Case 2:
        cmbStatus.Visible = False
        cmbTipo.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista2.SetFocus
        ProcCarregaListaHistorico
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Instrumentos = True Then
    USMsgBox ("Salve o instrumento antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    Permitido = False
    SSTab1.Tab = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposAfericao()
On Error GoTo tratar_erro

txtid_afericao = 0
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel1 = pubUsuario
txtAferido.Value = Date
txtOrgao = ""
txtProxima_Afericao.Value = Date
txtCertificado = ""
optSim.Value = False
optNao.Value = False
txt_Caminho = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaAfericao()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Afericao where ID_inst = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista1.ListItems
            .Add , , IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Aferido), "", Format(TBLISTA!Aferido, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Orgao), "", TBLISTA!Orgao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Proxima_afericao), "", Format(TBLISTA!Proxima_afericao, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
            If TBLISTA!Aprovado = True Then .Item(.Count).SubItems(5) = "Sim" Else .Item(.Count).SubItems(5) = "Não"
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

Private Sub ProcCarregaListaHistorico()
On Error GoTo tratar_erro

IDPlano = 0
Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select MD.IdPlano FROM Medicaodimensao_instrumentos MDI INNER JOIN Medicaodimensao MD ON MDI.idmedicao = MD.idmedicao where MDI.ID_inst = " & txtId & " order by MD.IdPlano", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        If IDPlano <> TBLISTA!IDPlano Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Medicao where IdPlano = " & TBLISTA!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                With Lista2.ListItems
                    .Add , , TBAbrir!IDPlano
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Peca), "", TBAbrir!Peca)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Inspetor), "", TBAbrir!Inspetor)
                End With
            End If
            TBAbrir.Close
        End If
        IDPlano = TBLISTA!IDPlano
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

Private Sub txtDPA_Change()
On Error GoTo tratar_erro
Dim DPA As Double

If IsNumeric(txtDPA.Text) Then

DPA = txtDPA.Text
txtINT.Text = DPA * 2

End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEMI_Change()
On Error GoTo tratar_erro
Dim varERM As Double
Dim varEMI As Double
Dim varINM As Double
Dim varEIM As Double

'=SE(J3=0;" ";(K3+L3))

If IsNumeric(txtERM.Text) Then
varERM = txtERM.Text
        If varERM = 0 Then
        txtEIM.Text = "0,00"
    Else
    If IsNumeric(txtEMI.Text) = True Then
        varEMI = IIf(txtEMI.Text <> "", txtEMI.Text, 0)
        varINM = IIf(txtINM.Text <> "", txtINM.Text, 0)
        varEIM = varEMI + varINM
        txtEIM.Text = varEIM
    End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtERM_Change()
On Error GoTo tratar_erro
Dim varERM As Double
Dim varEMI As Double
Dim varINM As Double
Dim varEIM As Double

'=SE(J3=0;" ";(K3+L3))

If IsNumeric(txtERM.Text) Then
varERM = txtERM.Text
        If varERM = 0 Then
        txtEIM.Text = "0,00"
    Else
        varEMI = IIf(txtEMI.Text <> "", txtEMI.Text, 0)
        varINM = IIf(txtINM.Text <> "", txtINM.Text, 0)
        varEIM = varEMI + varINM
        txtEIM.Text = varEIM
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtINM_Change()
On Error GoTo tratar_erro
Dim varERM As Double
Dim varEMI As Double
Dim varINM As Double
Dim varEIM As Double

'=SE(J3=0;" ";(K3+L3))

If IsNumeric(txtERM.Text) Then
varERM = txtERM.Text
        If varERM = 0 Then
        txtEIM.Text = "0,00"
    Else
        varEMI = IIf(txtEMI.Text <> "", txtEMI.Text, 0)
        varINM = IIf(txtINM.Text <> "", txtINM.Text, 0)
        varEIM = varEMI + varINM
        txtEIM.Text = varEIM
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtINT_Change()
On Error GoTo tratar_erro
Dim varINT As Double

If IsNumeric(txtINT.Text) Then

varINT = txtINT.Text
txtERM.Text = varINT / 3

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
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcStatus
    Case 9: procAtualiza
    'Case 11: ProcAjuda
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
    Case 1: ProcNovo_Afericao
    Case 2: ProcSalvar_Afericao
    Case 3: ProcExcluir_Afericao
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcAtualiza_Afericao
    'Case 9: ProcAjuda
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
    Case 1: ProcImprimir
    Case 2: ProcAnterior
    Case 3: ProcProximo
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=rQxAohwaWxQ&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=57&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
