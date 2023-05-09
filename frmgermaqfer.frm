VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGermaqfer 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Postos de trabalho"
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   75
      TabIndex        =   84
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   61
      Top             =   30
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frmgermaqfer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Listamaquina"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Framemaquina"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtidmaquina"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Turnos"
      TabPicture(1)   =   "frmgermaqfer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Txt_setfocus"
      Tab(1).Control(1)=   "USTreeView1"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "Framehoras"
      Tab(1).Control(4)=   "ImageList1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Instruções de trabalho"
      TabPicture(2)   =   "frmgermaqfer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "USToolBar3"
      Tab(2).Control(1)=   "Lista"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(3)=   "Txt_ID_instrucao"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Acessórios"
      TabPicture(3)   =   "frmgermaqfer.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "USToolBar4"
      Tab(3).Control(1)=   "Lista_acess"
      Tab(3).Control(2)=   "Txt_ID_acess"
      Tab(3).Control(3)=   "Frame4"
      Tab(3).ControlCount=   4
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
         Height          =   855
         Left            =   -74925
         TabIndex        =   95
         Top             =   1320
         Width           =   15195
         Begin VB.ComboBox Cmb_operacao_exec_acess 
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
            ItemData        =   "frmgermaqfer.frx":0070
            Left            =   12180
            List            =   "frmgermaqfer.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   47
            ToolTipText     =   "Operação para preparação."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox Txt_valor_exec_acess 
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
            Left            =   13650
            MaxLength       =   50
            TabIndex        =   48
            ToolTipText     =   "Valor para preparação."
            Top             =   390
            Width           =   1365
         End
         Begin VB.TextBox Txt_ID_produto_acess 
            Height          =   315
            Left            =   180
            TabIndex        =   100
            Text            =   "0"
            Top             =   390
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CommandButton Cmd_localizar_acess 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8910
            Picture         =   "frmgermaqfer.frx":008F
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Localizar acessórios."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox Txt_valor_prep_acess 
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
            Left            =   10800
            MaxLength       =   50
            TabIndex        =   46
            ToolTipText     =   "Valor para preparação."
            Top             =   390
            Width           =   1365
         End
         Begin VB.ComboBox Cmb_operacao_prep_acess 
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
            ItemData        =   "frmgermaqfer.frx":0191
            Left            =   9330
            List            =   "frmgermaqfer.frx":019B
            Style           =   2  'Dropdown List
            TabIndex        =   45
            ToolTipText     =   "Operação para preparação."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox Txt_descricao_acess 
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   390
            Width           =   7095
         End
         Begin VB.TextBox Txt_codigo_int_acess 
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
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. interno*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   525
            TabIndex        =   103
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Operação exec.*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   12285
            TabIndex        =   102
            Top             =   180
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor exec.*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   13882
            TabIndex        =   101
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor prep.*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   11040
            TabIndex        =   99
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Operação prep.*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   9442
            TabIndex        =   98
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   4995
            TabIndex        =   97
            Top             =   180
            Width           =   780
         End
      End
      Begin VB.TextBox Txt_ID_acess 
         Height          =   315
         Left            =   -71310
         TabIndex        =   94
         Text            =   "0"
         Top             =   4590
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtidmaquina 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Height          =   335
         Left            =   3270
         MaxLength       =   50
         MouseIcon       =   "frmgermaqfer.frx":01B0
         MousePointer    =   99  'Custom
         TabIndex        =   83
         Text            =   "0"
         ToolTipText     =   "Código da máquina."
         Top             =   5520
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -69030
         Top             =   4110
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   10
         ImageHeight     =   9
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgermaqfer.frx":04BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgermaqfer.frx":0824
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgermaqfer.frx":08A1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   76
         Top             =   9060
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
            ItemData        =   "frmgermaqfer.frx":0C4A
            Left            =   6960
            List            =   "frmgermaqfer.frx":0C5A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   187
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
            TabIndex        =   18
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
            TabIndex        =   16
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   22
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmgermaqfer.frx":0C8A
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
            TabIndex        =   21
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmgermaqfer.frx":442E
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
            TabIndex        =   19
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
            TabIndex        =   20
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmgermaqfer.frx":7F37
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
            TabIndex        =   23
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmgermaqfer.frx":C026
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
            Left            =   3360
            TabIndex        =   105
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   5610
            TabIndex        =   90
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   79
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
            TabIndex        =   78
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   2040
            TabIndex        =   77
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.TextBox Txt_ID_instrucao 
         Height          =   315
         Left            =   -71310
         TabIndex        =   66
         Text            =   "0"
         Top             =   4590
         Visible         =   0   'False
         Width           =   825
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
         Height          =   1125
         Left            =   -74925
         TabIndex        =   60
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_instrucao 
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
            Height          =   585
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            ToolTipText     =   "Instruções de trabalho."
            Top             =   390
            Width           =   14850
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instruções de trabalho*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   6743
            TabIndex        =   64
            Top             =   180
            Width           =   1725
         End
      End
      Begin VB.Frame Framehoras 
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
         Height          =   1455
         Left            =   -74925
         TabIndex        =   55
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtHora_extra 
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
            Left            =   13935
            TabIndex        =   39
            ToolTipText     =   "Percentual hora extra."
            Top             =   975
            Width           =   1065
         End
         Begin VB.TextBox txtIntervalo 
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
            Left            =   10620
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Total de horas do intervalo."
            Top             =   975
            Width           =   1065
         End
         Begin VB.TextBox txttotal1 
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
            Left            =   11715
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Total de horas do turno com intervalo."
            Top             =   975
            Width           =   1065
         End
         Begin VB.TextBox txthoras 
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
            Left            =   12795
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Total de horas disponíveis por dia."
            Top             =   975
            Width           =   1125
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   405
            Left            =   8505
            TabIndex        =   74
            Top             =   975
            Width           =   2175
            Begin MSComCtl2.DTPicker mskInicio_intervalo 
               Height          =   330
               Left            =   0
               TabIndex        =   34
               ToolTipText     =   "Início intervalo."
               Top             =   0
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
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
               CalendarTitleBackColor=   14737632
               CalendarTitleForeColor=   0
               CalendarTrailingForeColor=   12632256
               Format          =   198639618
               CurrentDate     =   39055
            End
            Begin MSComCtl2.DTPicker mskFinal_intervalo 
               Height          =   330
               Left            =   1050
               TabIndex        =   35
               ToolTipText     =   "Final do intervalo."
               Top             =   0
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
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
               CalendarTitleBackColor=   14737632
               CalendarTitleForeColor=   0
               CalendarTrailingForeColor=   12632256
               Format          =   198639618
               CurrentDate     =   39055
            End
         End
         Begin VB.CheckBox chkIntervalo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Intervalo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7500
            TabIndex        =   33
            Top             =   1035
            Width           =   960
         End
         Begin VB.TextBox txtData1 
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
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1575
         End
         Begin VB.TextBox txtResponsavel1 
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
            Left            =   1765
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   9245
         End
         Begin VB.TextBox txtStatus1 
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
            Left            =   11025
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   3975
         End
         Begin VB.TextBox txtturno 
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
            Left            =   6210
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Total de horas do turno sem intervalo."
            Top             =   975
            Width           =   1125
         End
         Begin VB.ComboBox cmbdia 
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
            ItemData        =   "frmgermaqfer.frx":F8B2
            Left            =   180
            List            =   "frmgermaqfer.frx":F8CB
            Style           =   2  'Dropdown List
            TabIndex        =   27
            ToolTipText     =   "Dia da semana."
            Top             =   975
            Width           =   2055
         End
         Begin VB.ComboBox cmbturno 
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
            ItemData        =   "frmgermaqfer.frx":F907
            Left            =   2250
            List            =   "frmgermaqfer.frx":F917
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Turno."
            Top             =   975
            Width           =   795
         End
         Begin MSComCtl2.DTPicker mskinicio 
            Height          =   330
            Left            =   3060
            TabIndex        =   29
            ToolTipText     =   "Início."
            Top             =   975
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
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
            CalendarTitleBackColor=   14737632
            CalendarTitleForeColor=   0
            CalendarTrailingForeColor=   12632256
            Format          =   198639618
            CurrentDate     =   39055
         End
         Begin MSComCtl2.DTPicker mskfinal 
            Height          =   330
            Left            =   5160
            TabIndex        =   31
            ToolTipText     =   "Final."
            Top             =   975
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
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
            CalendarTitleBackColor=   14737632
            CalendarTitleForeColor=   0
            CalendarTrailingForeColor=   12632256
            Format          =   198639618
            CurrentDate     =   39055
         End
         Begin MSComCtl2.DTPicker Msk_margem_ap 
            Height          =   330
            Left            =   4110
            TabIndex        =   30
            ToolTipText     =   "Margem para início do apontamento."
            Top             =   975
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
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
            CalendarTitleBackColor=   14737632
            CalendarTitleForeColor=   0
            CalendarTrailingForeColor=   12632256
            Format          =   198639618
            CurrentDate     =   39055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hora ex. (%)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   13987
            TabIndex        =   91
            Top             =   780
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Margem ap."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   4200
            TabIndex        =   75
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início inter."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   8655
            TabIndex        =   73
            Top             =   780
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Final inter."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   9675
            TabIndex        =   72
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   795
            TabIndex        =   71
            Top             =   195
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   5930
            TabIndex        =   70
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hs/dia"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   13132
            TabIndex        =   69
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Total turno"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   11850
            TabIndex        =   68
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Status"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   12780
            TabIndex        =   67
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   3345
            TabIndex        =   63
            Top             =   780
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Final*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   5467
            TabIndex        =   62
            Top             =   780
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hs/turno"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   6457
            TabIndex        =   59
            Top             =   780
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hs/intervalo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   10717
            TabIndex        =   58
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dia da semana*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   637
            TabIndex        =   57
            Top             =   780
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Turno*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   2392
            TabIndex        =   56
            Top             =   780
            Width           =   510
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   7215
         Left            =   -74925
         TabIndex        =   41
         Top             =   2460
         Width           =   15195
         _ExtentX        =   26802
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
            Text            =   "Instruções de trabalho"
            Object.Width           =   25585
         EndProperty
      End
      Begin VB.Frame Framemaquina 
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
         TabIndex        =   50
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox txtCustoHora_Setup 
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
            Left            =   11610
            MaxLength       =   50
            TabIndex        =   8
            ToolTipText     =   "Custo por hora de preparação."
            Top             =   990
            Width           =   1635
         End
         Begin VB.Frame Frame1 
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
            Height          =   1035
            Left            =   11910
            TabIndex        =   88
            Top             =   2100
            Width           =   3285
            Begin VB.CheckBox Chk_insp_final 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inspeção final?"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   180
               TabIndex        =   14
               Top             =   720
               Width           =   2925
            End
            Begin VB.CheckBox Optcustos 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Agrega custo/eficiência na ordem?"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   180
               TabIndex        =   12
               Top             =   180
               Width           =   2925
            End
            Begin VB.CheckBox Chk_liberada 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Liberada para apontamento?"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   180
               TabIndex        =   13
               Top             =   450
               Width           =   2925
            End
         End
         Begin VB.TextBox Txt_status 
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
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   3975
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   7485
         End
         Begin VB.TextBox Txt_data 
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
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   1575
         End
         Begin VB.TextBox txtGrupo 
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
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Grupo."
            Top             =   990
            Width           =   3615
         End
         Begin VB.CommandButton cmdGrupo 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3810
            Picture         =   "frmgermaqfer.frx":F927
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Localizar grupos."
            Top             =   990
            Width           =   315
         End
         Begin VB.CommandButton cmdSetor 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   11160
            Picture         =   "frmgermaqfer.frx":FA29
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Localizar setores."
            Top             =   990
            Width           =   315
         End
         Begin VB.TextBox txtMaquina 
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
            Left            =   13260
            MaxLength       =   50
            TabIndex        =   3
            ToolTipText     =   "Código da máquina."
            Top             =   390
            Width           =   1725
         End
         Begin VB.TextBox txtDescM 
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
            MaxLength       =   255
            TabIndex        =   10
            ToolTipText     =   "Descrição."
            Top             =   1620
            Width           =   14805
         End
         Begin VB.TextBox txtCustoHora 
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
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   "Custo por hora de execução."
            Top             =   990
            Width           =   1725
         End
         Begin VB.TextBox txtdados 
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
            Height          =   855
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Dados técnicos."
            Top             =   2250
            Width           =   11625
         End
         Begin VB.TextBox Txt_centro_de_custo 
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
            Left            =   4260
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Setor."
            Top             =   990
            Width           =   6885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Custo hora prep.*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   11767
            TabIndex        =   93
            Top             =   780
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Custo hora exec.*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   13455
            TabIndex        =   92
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Status"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   11025
            TabIndex        =   87
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   5055
            TabIndex        =   86
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   795
            TabIndex        =   85
            Top             =   195
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Grupo*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   1725
            TabIndex        =   65
            Top             =   780
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código*"
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
            Left            =   13785
            TabIndex        =   54
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição*"
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
            Index           =   1
            Left            =   7117
            TabIndex        =   53
            Top             =   1410
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dados técnicos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   5452
            TabIndex        =   52
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   7455
            TabIndex        =   51
            Top             =   780
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView Listamaquina 
         Height          =   4455
         Left            =   60
         TabIndex        =   15
         Top             =   4590
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   7858
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
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   10237
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Agrega custo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Liberada"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   80
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
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonToolTipText8=   "Copiar (F7)"
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
         ButtonWidth8    =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft9     =   347
         ButtonTop9      =   2
         ButtonWidth9    =   39
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Custo/eficiência"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Agrega custo/eficiência na ordem (F9)"
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
         ButtonLeft10    =   388
         ButtonTop10     =   2
         ButtonWidth10   =   84
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Liberar"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Liberar posto para apontamento (F10)"
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
         ButtonLeft11    =   474
         ButtonTop11     =   2
         ButtonWidth11   =   41
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   517
         ButtonTop12     =   2
         ButtonWidth12   =   50
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
         ButtonLeft13    =   569
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft14    =   573
         ButtonTop14     =   2
         ButtonWidth14   =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft15    =   611
         ButtonTop15     =   2
         ButtonWidth15   =   26
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
         ButtonLeft16    =   639
         ButtonTop16     =   2
         ButtonWidth16   =   24
         ButtonHeight16  =   24
         ButtonUseMaskColor16=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   12900
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmgermaqfer.frx":FB2B
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   81
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
         ButtonCaption7  =   "Copiar"
         ButtonEnabled7  =   0   'False
         ButtonToolTipText7=   "Copiar (F7)"
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
         ButtonWidth7    =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft8     =   309
         ButtonTop8      =   2
         ButtonWidth8    =   39
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
         ButtonLeft9     =   350
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
         ButtonLeft10    =   354
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
         ButtonLeft11    =   392
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
         ButtonLeft12    =   420
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   11160
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmgermaqfer.frx":190BF
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   82
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
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   11310
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmgermaqfer.frx":1F90D
            Count           =   1
         End
      End
      Begin DrawSuite2022.USTreeView USTreeView1 
         Height          =   6885
         Left            =   -74910
         TabIndex        =   89
         Top             =   2790
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12144
         BorderColor     =   12500670
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Theme           =   1
      End
      Begin MSComctlLib.ListView Lista_acess 
         Height          =   7485
         Left            =   -74925
         TabIndex        =   49
         Top             =   2190
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13203
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
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   14296
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Operação prep."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Valor prep."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Operação exec."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Valor exec."
            Object.Width           =   2117
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74925
         TabIndex        =   96
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
         Begin DrawSuite2022.USImageList USImageList4 
            Left            =   11310
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmgermaqfer.frx":24CF1
            Count           =   1
         End
      End
      Begin VB.TextBox Txt_setfocus 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Height          =   335
         Left            =   -73620
         MaxLength       =   50
         MouseIcon       =   "frmgermaqfer.frx":2A0D5
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   4020
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmGermaqfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_maquina                    As Boolean
Public Novo_maquina1                As Boolean
Dim Novo_maquina2                   As Boolean
Dim Novo_maquina3                   As Boolean
Public Sql_maquina_Localizar        As String 'OK
Dim CodMaquina                      As String  'OK
Dim TBLISTA_Posto_Trabalho          As ADODB.Recordset 'OK
Public FormulaRel_Posto_Trabalho    As String 'OK

Private Sub chkIntervalo_Click()
On Error GoTo tratar_erro

If chkIntervalo.Value = 1 Then
    Frame3.Enabled = True
    mskInicio_intervalo.SetFocus
Else
    mskInicio_intervalo = "00:00:00"
    mskFinal_intervalo = "00:00:00"
    Frame3.Enabled = False
End If
ProcCalculaTempos
ProcCalculaTemposIntervalo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With ListaMaquina
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    Select Case Cmb_opcao_lista
        Case "Excluir":
            .ButtonState(4) = 0
            .ButtonState(9) = 5
            .ButtonState(10) = 5
            .ButtonState(11) = 5
        Case "Status":
            .ButtonState(4) = 5
            .ButtonState(9) = 0
            .ButtonState(10) = 5
            .ButtonState(11) = 5
        Case "Custo/eficiência":
            .ButtonState(4) = 5
            .ButtonState(9) = 5
            .ButtonState(10) = 0
            .ButtonState(11) = 5
        Case "Liberar":
            .ButtonState(4) = 5
            .ButtonState(9) = 5
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

Private Sub cmbdia_Click()
On Error GoTo tratar_erro

ProcCarregaDadosTurno

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbturno_Click()
On Error GoTo tratar_erro

ProcCarregaDadosTurno

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDadosTurno()
On Error GoTo tratar_erro

If txtmaquina.Text <> "" And cmbdia.Text <> "" And cmbturno.Text <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from cadmaqturnos where maquina = '" & txtmaquina & "' and diasemana = '" & cmbdia & "' and Turno = " & cmbturno, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        mskinicio.Value = IIf(IsNull(TBAbrir!Inicioturno), "00:00:00", Left(TBAbrir!Inicioturno, 8))
        Msk_margem_ap.Value = IIf(IsNull(TBAbrir!Margem_inicio_ap), "00:00:00", Left(TBAbrir!Margem_inicio_ap, 8))
        mskfinal.Value = IIf(IsNull(TBAbrir!finalturno), "00:00:00", Left(TBAbrir!finalturno, 8))
        If TBAbrir!intervalo <> "" And Left(TBAbrir!intervalo, 8) <> "00:00:00" Then chkIntervalo.Value = 1 Else chkIntervalo.Value = 0
        txtIntervalo.Text = IIf(IsNull(TBAbrir!intervalo), "00:00:00", Left(TBAbrir!intervalo, 8))
        mskInicio_intervalo.Value = IIf(IsNull(TBAbrir!Inicio_intervalo), "00:00:00", Left(TBAbrir!Inicio_intervalo, 8))
        mskFinal_intervalo.Value = IIf(IsNull(TBAbrir!Final_intervalo), "00:00:00", Left(TBAbrir!Final_intervalo, 8))
        txtTurno.Text = Left(TBAbrir!Total, 8)
        txttotal1.Text = Left(TBAbrir!TotalTurno, 8)
        txtHora_extra.Text = IIf(IsNull(TBAbrir!Percentual_HoraExtra), "", TBAbrir!Percentual_HoraExtra)
        
        If TBAbrir!TotalDia <> "01/01/1900" Then
            Dataini = Right(TBAbrir!TotalDia, 8)
            ElapsedTime (Dataini)
            txthoras.Text = HoraTotal
        Else
            txthoras = "00:00:00"
        End If
        If D > 24 Then
            USMsgBox ("Atenção, total de horas disponíveis do dia está excedendo a 24 horas, corrija os turnos."), vbExclamation, "CAPRIND v5.0"
        End If
    Else
        mskinicio.Value = "00:00:00"
        Msk_margem_ap.Value = "00:00:00"
        mskfinal.Value = "00:00:00"
        chkIntervalo.Value = 0
        mskInicio_intervalo.Value = "00:00:00"
        mskFinal_intervalo.Value = "00:00:00"
        txtTurno.Text = "00:00:00"
        txtIntervalo.Text = "00:00:00"
        txttotal1.Text = "00:00:00"
        txthoras.Text = ""
        txtHora_extra.Text = ""
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloquearTurno()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmGermaqfer_bloqTurno.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirInstrucao()
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
                If USMsgBox("Deseja realmente excluir esta(s) instrução(ões) de trabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from CadMaquinas_instrucoes where ID = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "PCP/Postos de trabalho"
            Evento = "Excluir instrução de trabalho"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Código do posto de trabalho: " & txtmaquina
            Documento1 = "Instrução de trabalho: " & .ListItems.Item(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) instrução(ões) de trabalho antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Instrução(ões) de trabalho excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposInstrucoes
    ProcCarregaListaInstrucoes
    Frame2.Enabled = False
    Novo_maquina2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirAcess()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista_acess
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) acessório(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from CadMaquinas_acessorios where ID = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "PCP/Postos de trabalho"
            Evento = "Excluir instrução de trabalho"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Código do posto de trabalho: " & txtmaquina
            Documento1 = "Acessório: " & .ListItems.Item(InitFor).ListSubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) acessório(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Acessório(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposAcess
    ProcCarregaListaAcess
    Frame4.Enabled = False
    Novo_maquina3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoInstrucao()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Novo_maquina2 = True
ProcLimpaCamposInstrucoes
Frame2.Enabled = True
Txt_instrucao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoAcess()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Novo_maquina3 = True
ProcLimpaCamposAcess
Frame4.Enabled = True
Cmd_localizar_acess_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposInstrucoes()
On Error GoTo tratar_erro

Txt_ID_instrucao = 0
Txt_instrucao = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposAcess()
On Error GoTo tratar_erro

Txt_ID_acess = 0
Txt_ID_produto_acess = 0
Txt_codigo_int_acess = ""
Txt_descricao_acess = ""
Cmb_operacao_prep_acess.ListIndex = -1
Txt_valor_prep_acess = ""
Cmb_operacao_exec_acess.ListIndex = -1
Txt_valor_exec_acess = ""
CodigoLista3 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirTurno()
On Error GoTo tratar_erro

If Framehoras.Enabled = False Then
    USMsgBox ("Informe o turno antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_maquina1 = True Then
    USMsgBox ("Salve o turno antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excluir o turno " & cmbturno.Text & " de " & cmbdia.Text & " do posto de trabalho " & txtmaquina.Text & ".", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select * from  CadMaqturnos where Maquina = '" & txtmaquina.Text & "' and diasemana = '" & cmbdia.Text & "' and turno = " & cmbturno.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        '==================================
        Modulo = "PCP/Postos de trabalho"
        Evento = "Excluir turno"
        ID_documento = TBMaquinas!CODIGO
        Documento = "Código do posto de trabalho: " & txtmaquina.Text
        Documento1 = "Dia da semana: " & cmbdia.Text & " - Truno: " & cmbturno.Text
        ProcGravaEvento
        '==================================
        TBMaquinas.Delete
        USMsgBox ("Turno excluído com sucesso."), vbInformation, "CAPRIND v5.0"
        ProcRecalculaTempoTotalDia
        ProcLimpaCamposTurno
        Novo_maquina1 = False
        Framehoras.Enabled = False
        ProcCarregaTurnos
    End If
End If
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a alterar neste formulário."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If txtmaquina = "" Then
    USMsgBox ("Informe o posto de trabalho antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_maquina = True Then
    USMsgBox ("Salve o posto de trabalho antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente copiar os dados deste posto de trabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    CodMaquina = InputBox("Favor informar o novo código do posto de trabalho.")
    If CodMaquina = "" Then Exit Sub
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select idmaquina from cadmaquinas where maquina = '" & CodMaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        USMsgBox ("Código do posto de trabalho já existente, favor alterar."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBGravar.Close
    
    'Dados principais
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select * from cadmaquinas where idmaquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from cadmaquinas", Conexao, adOpenKeyset, adLockOptimistic
        TBAbrir.AddNew
        TBAbrir!Data = Date
        TBAbrir!Responsavel = pubUsuario
        If TBMaquinas!custos = True Then TBAbrir!custos = True Else TBAbrir!custos = False
        TBAbrir!Liberada = True
        TBAbrir!maquina = CodMaquina
        TBAbrir!Grupo = TBMaquinas!Grupo
        TBAbrir!Setor = TBMaquinas!Setor
        TBAbrir!PrecoHora = TBMaquinas!PrecoHora
        TBAbrir!PrecoHora_Setup = TBMaquinas!PrecoHora_Setup
        TBAbrir!Descricao = TBMaquinas!Descricao
        TBAbrir!caracteristicas = TBMaquinas!caracteristicas
        TBAbrir.Update
    
        'Turnos
        Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select * from CadmaqTurnos where maquina = '" & txtmaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAfericao.EOF = False Then
            Do While TBAfericao.EOF = False
                Set TBAliquota = CreateObject("adodb.recordset")
                TBAliquota.Open "Select * from CadmaqTurnos", Conexao, adOpenKeyset, adLockOptimistic
                TBAliquota.AddNew
                TBAliquota!maquina = CodMaquina
                TBAliquota!Data = Date
                TBAliquota!Responsavel = pubUsuario
                TBAliquota!Diasemana = TBAfericao!Diasemana
                TBAliquota!Turno = TBAfericao!Turno
                TBAliquota!Inicioturno = TBAfericao!Inicioturno
                TBAliquota!Margem_inicio_ap = TBAfericao!Margem_inicio_ap
                TBAliquota!finalturno = TBAfericao!finalturno
                TBAliquota!Total = TBAfericao!Total
                TBAliquota!intervalo = TBAfericao!intervalo
                TBAliquota!Inicio_intervalo = TBAfericao!Inicio_intervalo
                TBAliquota!Final_intervalo = TBAfericao!Final_intervalo
                TBAliquota!TotalTurno = TBAfericao!TotalTurno
                TBAliquota!TotalDia = TBAfericao!TotalDia
                TBAliquota!Percentual_HoraExtra = TBAfericao!Percentual_HoraExtra
                TBAliquota.Update
                TBAliquota.Close
                TBAfericao.MoveNext
            Loop
        End If
        TBAfericao.Close
        
        'Instruções de trabalho
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select * from CadMaquinas_instrucoes where ID_maquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = False Then
            Do While TBAcessos.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from CadMaquinas_instrucoes", Conexao, adOpenKeyset, adLockOptimistic
                TBGravar.AddNew
                TBGravar!ID_Maquina = TBAbrir!IDMaquina
                TBGravar!Instrucao = TBAcessos!Instrucao
                TBGravar.Update
                TBGravar.Close
                TBAcessos.MoveNext
            Loop
        End If
        TBAcessos.Close
        
        'Acessórios
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select * from CadMaquinas_acessorios where ID_posto = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = False Then
            Do While TBAcessos.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from CadMaquinas_acessorios", Conexao, adOpenKeyset, adLockOptimistic
                TBGravar.AddNew
                TBGravar!ID_posto = TBAbrir!IDMaquina
                TBGravar!ID_Produto = TBAcessos!ID_Produto
                TBGravar!Operacao_prep = TBAcessos!Operacao_prep
                TBGravar!Valor_prep = TBAcessos!Valor_prep
                TBGravar!Operacao_exec = TBAcessos!Operacao_exec
                TBGravar!Valor_exec = TBAcessos!Valor_exec
                TBGravar.Update
                TBGravar.Close
                TBAcessos.MoveNext
            Loop
        End If
        TBAcessos.Close
        TBAbrir.Close
        
        USMsgBox ("Posto de trabalho copiado com sucesso."), vbInformation, "CAPRIND v5.0"
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * FROM Cadmaquinas WHERE maquina = '" & CodMaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            ProcLimpaCampos
            ProcPuxaDados
        End If
        TBAbrir.Close
        ProcAtualizaListaPT (1)
        '==================================
        Modulo = "PCP/Postos de trabalho"
        Evento = "Novo"
        ID_documento = txtIDmaquina
        Documento = "Código do posto de trabalho: " & txtmaquina.Text
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
    TBMaquinas.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarTurno()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a alterar neste formulário."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
frmGermaqfer_MenuCopiar.Show 1

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
With ListaMaquina
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) posto(s) de trabalho antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmGermaqfer_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCustoEficiencia()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaMaquina
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o campo agrega custo/eficiência na ordem neste(s) posto(s) de trabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "UPDATE cadmaquinas set custos = 1 where IDmaquina = " & .ListItems.Item(InitFor)
            '==================================
            Modulo = "PCP/Postos de trabalho"
            Evento = "Agregar custo/eficiência na ordem"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Código do posto de trabalho: " & .ListItems.Item(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) posto(s) de trabalho antes de alterar o campo agrega custo/eficiência na ordem."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Posto(s) de trabalho alterado(s) com sucesso."), vbExclamation, "CAPRIND v5.0"
    ProcAtualizaListaPT (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CadMaquinas where IDMaquina = " & IIf(txtIDmaquina = "", 0, txtIDmaquina), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcPuxaDados
    End If
    TBAbrir.Close
End If

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
Permitido = False
With ListaMaquina
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente liberar este(s) posto(s) de trabalho para apontamento?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "UPDATE cadmaquinas set Liberada = 1 where IDmaquina = " & .ListItems.Item(InitFor)
            '==================================
            Modulo = "PCP/Postos de trabalho"
            Evento = "Liberar"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Código do posto de trabalho: " & .ListItems.Item(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) posto(s) de trabalho antes de liberar para apontamento."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Posto(s) de trabalho liberados(s) com sucesso."), vbExclamation, "CAPRIND v5.0"
    ProcAtualizaListaPT (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CadMaquinas where IDMaquina = " & IIf(txtIDmaquina = "", 0, txtIDmaquina), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ProcPuxaDados
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTurno()
On Error GoTo tratar_erro

Dataini = "00:00:00"
DataFim = "00:00:00"
Acao = "salvar"
If Framehoras.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If cmbdia.Text = "" Then
    NomeCampo = "o dia da semana"
    ProcVerificaAcao
    cmbdia.SetFocus
    Exit Sub
End If
If cmbturno.Text = "" Then
    NomeCampo = "o turno"
    ProcVerificaAcao
    cmbturno.SetFocus
    Exit Sub
End If
If mskinicio.Value = "00:00:00" Then
    NomeCampo = "o início"
    ProcVerificaAcao
    mskinicio.SetFocus
    Exit Sub
End If
If Msk_margem_ap.Value <> "00:00:00" Then
    ElapsedTime (Msk_margem_ap)
    If s > 3600 Then
        USMsgBox ("A margem para apontamento não pode ser maior que uma hora."), vbExclamation, "CAPRIND v5.0"
        Msk_margem_ap.SetFocus
        Exit Sub
    End If
End If
If mskfinal.Value = "00:00:00" Then
    NomeCampo = "o final"
    ProcVerificaAcao
    mskfinal.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from cadmaqturnos where maquina = '" & txtmaquina.Text & "' and diasemana = '" & cmbdia.Text & "' and turno = " & cmbturno.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    USMsgBox ("Novo turno cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo turno"
    TBGravar!Bloqueado = False
Else
    Evento = "Alterar turno"
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
End If
TBGravar!maquina = txtmaquina.Text
TBGravar!Data = IIf(txtData1.Text = "", Date, txtData1.Text)
TBGravar!Responsavel = IIf(txtResponsavel1.Text = "", pubUsuario, txtResponsavel1)
TBGravar!Diasemana = cmbdia.Text
TBGravar!Turno = cmbturno.Text
TBGravar!Inicioturno = mskinicio.Value
TBGravar!Margem_inicio_ap = Msk_margem_ap.Value
TBGravar!finalturno = mskfinal.Value
TBGravar!Total = txtTurno.Text
TBGravar!intervalo = txtIntervalo.Text
TBGravar!Inicio_intervalo = mskInicio_intervalo.Value
TBGravar!Final_intervalo = mskFinal_intervalo.Value
TBGravar!TotalTurno = txttotal1.Text
TBGravar!Percentual_HoraExtra = IIf(txtHora_extra.Text = "", 0, txtHora_extra.Text)
TBGravar.Update

'==================================
Modulo = "PCP/Postos de trabalho"
ID_documento = TBGravar!CODIGO
Documento = "Código do posto de trabalho: " & txtmaquina.Text
Documento1 = "Dia da semana: " & cmbdia.Text & " - Truno: " & cmbturno.Text
ProcGravaEvento
'==================================
TBGravar.Close
ProcRecalculaTempoTotalDia
Novo_maquina1 = False
ProcCarregaTurnos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRecalculaTempoTotalDia()
On Error GoTo tratar_erro

DataFim = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from cadmaqturnos where maquina = '" & txtmaquina.Text & "' and diasemana = '" & cmbdia.Text & "' and bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
Do While TBAbrir.EOF = False
    Dataini = Left(TBAbrir!TotalTurno, 8)
    DataFim = Format(DataFim + Dataini, "hh:mm:ss")
    TBAbrir.MoveNext
Loop
TBAbrir.Close
Conexao.Execute "Update cadmaqturnos Set TotalDia = '" & Format(DataFim, "hh:mm:ss") & "' where maquina = '" & txtmaquina.Text & "' and diasemana = '" & cmbdia.Text & "'"
ElapsedTime (DataFim)
txthoras = HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_acess_Click()
On Error GoTo tratar_erro
     
CadMaquinas = True
Estoque_entrada = False
frmEstoque_fisico_item.Show 1
                
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdGrupo_Click()
On Error GoTo tratar_erro
     
frmGermaqfer_grupo.Show 1
                
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoturno()
On Error GoTo tratar_erro

ProcLimpaCamposTurno
Framehoras.Enabled = True
cmbdia.SetFocus
Novo_maquina1 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarInstrucao()
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
If Txt_instrucao = "" Then
    NomeCampo = "a instrução de trabalho"
    ProcVerificaAcao
    Txt_instrucao.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CadMaquinas_instrucoes where ID = " & Txt_ID_instrucao, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_Maquina = txtIDmaquina
TBGravar!Instrucao = Txt_instrucao
TBGravar.Update
Txt_ID_instrucao = TBGravar!ID
TBGravar.Close
ProcCarregaListaInstrucoes
If Novo_maquina2 = True Then
    USMsgBox ("Nova instrução de trabalho cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova instrução de trabalho"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar instrução de trabalho"
    If CodigoLista2 <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista2)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "PCP/Postos de trabalho"
ID_documento = Txt_ID_instrucao
Documento = "Código do posto de trabalho: " & txtmaquina
Documento1 = "Instrução de trabalho: " & Txt_instrucao
ProcGravaEvento
'==================================
Novo_maquina2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarAcess()
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
If Txt_codigo_int_acess = "" Then
    NomeCampo = "o acessório"
    ProcVerificaAcao
    Cmd_localizar_acess_Click
    Exit Sub
End If
If Cmb_operacao_prep_acess = "" Then
    NomeCampo = "a operação para preparação"
    ProcVerificaAcao
    Cmb_operacao_prep_acess.SetFocus
    Exit Sub
End If
If Txt_valor_prep_acess = "" Then
    NomeCampo = "o valor para preparação"
    ProcVerificaAcao
    Txt_valor_prep_acess.SetFocus
    Exit Sub
End If
If Cmb_operacao_exec_acess = "" Then
    NomeCampo = "a operação para execução"
    ProcVerificaAcao
    Cmb_operacao_exec_acess.SetFocus
    Exit Sub
End If
If Txt_valor_exec_acess = "" Then
    NomeCampo = "o valor para execução"
    ProcVerificaAcao
    Txt_valor_exec_acess.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from CadMaquinas_acessorios where ID = " & Txt_ID_acess, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_posto = txtIDmaquina
TBGravar!ID_Produto = Txt_ID_produto_acess
TBGravar!Operacao_prep = IIf(Cmb_operacao_prep_acess = "Somar", 1, 2)
TBGravar!Valor_prep = Txt_valor_prep_acess
TBGravar!Operacao_exec = IIf(Cmb_operacao_exec_acess = "Somar", 1, 2)
TBGravar!Valor_exec = Txt_valor_exec_acess
TBGravar.Update
Txt_ID_acess = TBGravar!ID
TBGravar.Close
ProcCarregaListaAcess
If Novo_maquina3 = True Then
    USMsgBox ("Novo acessório cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Nova acessório"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar instrução de trabalho"
    If CodigoLista3 <> 0 And Lista_acess.ListItems.Count <> 0 Then
        Lista_acess.SelectedItem = Lista_acess.ListItems(CodigoLista3)
        Lista_acess.SetFocus
    End If
End If
'==================================
Modulo = "PCP/Postos de trabalho"
ID_documento = Txt_ID_acess
Documento = "Código do posto de trabalho: " & txtmaquina
Documento1 = "Acessório: " & Txt_descricao_acess
ProcGravaEvento
'==================================
Novo_maquina3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Posto_Trabalho.AbsolutePage <> 2 Then
    If TBLISTA_Posto_Trabalho.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Posto_Trabalho.PageCount - 1)
    Else
        TBLISTA_Posto_Trabalho.AbsolutePage = TBLISTA_Posto_Trabalho.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Posto_Trabalho.AbsolutePage)
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
    TBLISTA_Posto_Trabalho.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Posto_Trabalho.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Posto_Trabalho.AbsolutePage = 1
ProcExibePagina (TBLISTA_Posto_Trabalho.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Posto_Trabalho.AbsolutePage <> -3 Then
    If TBLISTA_Posto_Trabalho.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Posto_Trabalho.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Posto_Trabalho.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Posto_Trabalho.AbsolutePage = TBLISTA_Posto_Trabalho.PageCount
ProcExibePagina (TBLISTA_Posto_Trabalho.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSetor_Click()
On Error GoTo tratar_erro

CadMaquinas = True
Funcionario = False
Usuarios = False
Estoque_Local_Armazenamento = False
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
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcGravar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCopiar
            Case vbKeyF8: If Cmb_opcao_lista = "Status" Then ProcStatus
            Case vbKeyF9: If Cmb_opcao_lista = "Custo/eficiência" Then ProcCustoEficiencia
            Case vbKeyF10: If Cmb_opcao_lista = "Liberar" Then ProcLiberar
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoturno
            Case vbKeyF3: ProcGravarTurno
            Case vbKeyF4: ProcExcluirTurno
            Case vbKeyF7: ProcCopiarTurno
            Case vbKeyF8: ProcBloquearTurno
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoInstrucao
            Case vbKeyF3: ProcGravarInstrucao
            Case vbKeyF4: ProcExcluirInstrucao
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoAcess
            Case vbKeyF3: ProcGravarAcess
            Case vbKeyF4: ProcExcluirAcess
            'Case vbKeyF1: ProcAjuda
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

ProcCarregaToolBar1 Me, 15192, 16, True
ProcCarregaToolBar2 Me, 15192, 11, True
ProcCarregaToolBar3 Me, 15192, 9, True
ProcCarregaToolBar4 Me, 15192, 9, True

Formulario = "PCP/Postos de trabalho"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Cmb_opcao_lista = "Excluir"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362M" Then
    If USMsgBox("Deseja realmente atualizar a liberação dos postos de trabalho e se o posto de trabalho agrega custos\eficiencia nas ordens?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select maquina,evento,liberada,custos from cadmaquinas order by maquina", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            TBMaquinas.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBMaquinas.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBMaquinas.MoveFirst
            Do While TBMaquinas.EOF = False
                If IsNull(TBMaquinas!Evento) = False Then
                    Set TBCodigoDesc = CreateObject("adodb.recordset")
                    TBCodigoDesc.Open "Select * from CodigoDesc where codigo = " & TBMaquinas!Evento & " and Liberar_Posto = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCodigoDesc.EOF = False Then
                        TBMaquinas!Liberada = True
                    Else
                        TBMaquinas!Liberada = False
                    End If
                    TBCodigoDesc.Close
                Else
                    TBMaquinas!Liberada = True
                End If
                TBMaquinas.Update
                 'Altera na OS se tem custos
                If TBMaquinas!custos = True Then
                    Conexao.Execute "Update Ordemservico Set Custos = 'True' where maquina = '" & TBMaquinas!maquina & "'"
                Else
                    Conexao.Execute "Update Ordemservico Set Custos = 'False' where maquina = '" & TBMaquinas!maquina & "'"
                End If
                TBMaquinas.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBMaquinas.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "PCP/Postos de trabalho"
        Evento = "Atualizar"
        Documento = 0
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If ListaMaquina.ListItems.Count = 0 Then Exit Sub
frmGermaqfer_menu_impressao.Show 1

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
Framemaquina.Enabled = True
txtmaquina.SetFocus
Novo_maquina = True
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Framehoras.Enabled = False
Frame2.Enabled = False
Frame4.Enabled = False
ProcLimpaCamposTurno
ProcLimpaCamposInstrucoes
ProcLimpaCamposAcess
USTreeView1.Clear
Lista.ListItems.Clear
Lista_acess.ListItems.Clear
Novo_maquina1 = False
Novo_maquina2 = False
Novo_maquina3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
    
frmmaquina_Abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
  
If Novo_maquina = True Then
    If USMsgBox("O posto de trabalho ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
       ProcGravar
        If Novo_maquina = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_maquina1 = True Then
    If USMsgBox("O turno ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarTurno
        If Novo_maquina1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_maquina2 = True Then
    If USMsgBox("A instrução de trabalho ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravarInstrucao
        If Novo_maquina2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_maquina = False
Novo_maquina1 = False
Novo_maquina2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Framemaquina.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtmaquina.Text = "" Then
    NomeCampo = "o código"
    ProcVerificaAcao
    txtmaquina.SetFocus
    Exit Sub
End If
If txtGrupo.Text = "" Then
    NomeCampo = "o grupo"
    ProcVerificaAcao
    frmGermaqfer_grupo.Show 1
    Exit Sub
End If
If Txt_centro_de_custo.Text = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    CadMaquinas = True
    Funcionario = False
    Usuarios = False
    Estoque_Local_Armazenamento = False
    frmUsuarios_Setor.Show 1
    Exit Sub
End If
If txtCustoHora_Setup.Text = "" Then
    NomeCampo = "o custo hora por preparação"
    ProcVerificaAcao
    txtCustoHora_Setup.SetFocus
    Exit Sub
End If
If txtCustoHora.Text = "" Then
    NomeCampo = "o custo hora"
    ProcVerificaAcao
    txtCustoHora.SetFocus
    Exit Sub
End If
If txtDescM.Text = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescM.SetFocus
    Exit Sub
End If

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from cadmaquinas where IDMaquina = " & IIf(txtIDmaquina = "", 0, txtIDmaquina), Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    If txtmaquina <> TBMaquinas!maquina Then
        Conexao.Execute "Update CadmaqTurnos Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update CQ_NC_FABRICA Set maquina = '" & txtmaquina & "', Setor = '" & Txt_centro_de_custo & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update Fases Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update Manutencao Set IDMaquina = '" & txtmaquina & "' where IDMaquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update Ordemservico Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update Ordemservico_maq_utilizadas Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update ProducaoFases Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update ProducaoFases_Backup Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update ProducaoFases_Totalizacao Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update ProducaoFases_Totalizacao_Backup Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
        Conexao.Execute "Update Programas Set maquina = '" & txtmaquina & "' where maquina = '" & TBMaquinas!maquina & "'"
    End If
Else
    TBMaquinas.AddNew
    TBMaquinas!Bloqueado = False
    TBMaquinas!Liberada = True
End If
IDMaquina = IIf(IsNull(TBMaquinas!IDMaquina), 0, TBMaquinas!IDMaquina)
TBMaquinas!Data = IIf(Txt_data = "", Date, Txt_data)
TBMaquinas!Responsavel = IIf(Txt_responsavel = "", pubUsuario, Txt_responsavel)
If Optcustos.Value = 1 Then TBMaquinas!custos = True Else TBMaquinas!custos = False
If Chk_liberada.Value = 1 Then TBMaquinas!Liberada = True Else TBMaquinas!Liberada = False
If Chk_insp_final.Value = 1 Then TBMaquinas!Insp_final = True Else TBMaquinas!Insp_final = False
TBMaquinas!maquina = txtmaquina.Text
TBMaquinas!Grupo = txtGrupo
TBMaquinas!Setor = Txt_centro_de_custo.Text
TBMaquinas!PrecoHora = txtCustoHora.Text
TBMaquinas!PrecoHora_Setup = txtCustoHora_Setup.Text
TBMaquinas!Descricao = txtDescM.Text
TBMaquinas!caracteristicas = txtdados.Text
TBMaquinas.Update
txtIDmaquina.Text = TBMaquinas!IDMaquina
TBMaquinas.Close

If Novo_maquina = True Then
    USMsgBox "Novo posto de trabalho cadastrado com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_maquina_Localizar = "SELECT * from cadmaquinas where IDMaquina = " & txtIDmaquina
    ProcAtualizaListaPT (1)
Else
    USMsgBox "Alteração efetuada com sucesso.", vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizaListaPT (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And ListaMaquina.ListItems.Count <> 0 Then
       ListaMaquina.SelectedItem = ListaMaquina.ListItems(CodigoLista)
       ListaMaquina.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "PCP/Postos de trabalho"
    ID_documento = txtIDmaquina
    Documento = "Código do posto de trabalho: " & txtmaquina.Text
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_maquina = False
  
Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaListaPT(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListaMaquina.ListItems.Clear
If Sql_maquina_Localizar = "" Then Exit Sub
Set TBLISTA_Posto_Trabalho = CreateObject("adodb.recordset")
TBLISTA_Posto_Trabalho.Open Sql_maquina_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Posto_Trabalho.EOF = False Then ProcExibePagina (Pagina)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaMaquina.ListItems.Clear
TBLISTA_Posto_Trabalho.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Posto_Trabalho.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Posto_Trabalho.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Posto_Trabalho.RecordCount - IIf(Pagina > 1, (TBLISTA_Posto_Trabalho.PageSize * (Pagina - 1)), 0), TBLISTA_Posto_Trabalho.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Posto_Trabalho.EOF = False And (ContadorReg <= TamanhoPagina)
   With ListaMaquina.ListItems
        .Add , , TBLISTA_Posto_Trabalho!IDMaquina
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Posto_Trabalho!Data), "", Format(TBLISTA_Posto_Trabalho!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Posto_Trabalho!Responsavel), "", TBLISTA_Posto_Trabalho!Responsavel)
        .Item(.Count).SubItems(3) = TBLISTA_Posto_Trabalho!maquina
        .Item(.Count).SubItems(4) = TBLISTA_Posto_Trabalho!Descricao
        If TBLISTA_Posto_Trabalho!custos = True Then .Item(.Count).SubItems(5) = "SIM" Else .Item(.Count).SubItems(5) = "NÃO"
        If TBLISTA_Posto_Trabalho!Liberada = True Then .Item(.Count).SubItems(6) = "SIM" Else .Item(.Count).SubItems(6) = "NÃO"
        If TBLISTA_Posto_Trabalho!Bloqueado = True Then .Item(.Count).SubItems(7) = "Bloqueado" Else .Item(.Count).SubItems(7) = "Liberado"
    End With
    TBLISTA_Posto_Trabalho.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Posto_Trabalho.RecordCount
If TBLISTA_Posto_Trabalho.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Posto_Trabalho.PageCount
ElseIf TBLISTA_Posto_Trabalho.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Posto_Trabalho.PageCount & " de: " & TBLISTA_Posto_Trabalho.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Posto_Trabalho.AbsolutePage - 1 & " de: " & TBLISTA_Posto_Trabalho.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaInstrucoes()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CadMaquinas_instrucoes where ID_maquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    With Lista.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!Instrucao
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

Private Sub ProcCarregaListaAcess()
On Error GoTo tratar_erro

Lista_acess.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CMA.*, P.Desenho, P.Descricao from CadMaquinas_acessorios CMA INNER JOIN Projproduto P ON P.Codproduto = CMA.ID_produto where CMA.ID_posto = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    With Lista_acess.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            If TBLISTA!Operacao_prep = 1 Then .Item(.Count).SubItems(3) = "Somar" Else .Item(.Count).SubItems(3) = "Subtrair"
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Valor_prep), "", Format(TBLISTA!Valor_prep, "###,##0.00"))
            If TBLISTA!Operacao_exec = 1 Then .Item(.Count).SubItems(5) = "Somar" Else .Item(.Count).SubItems(5) = "Subtrair"
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Valor_exec), "", Format(TBLISTA!Valor_exec, "###,##0.00"))
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

Private Sub Lista_acess_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_acess
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                ProcVerificaRegistroUtilizadoSemMsg "Ferramentas", "ID_acessorio = " & .ListItems(InitFor)
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
    ProcOrdenaListView Lista_acess, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_acess_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_acess
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Mensagem = "Não é permitido excluir este acessório, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "Ferramentas", "ID_acessorio = " & .ListItems(InitFor), "Engenharia/Processos"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_acess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_acess.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select CMA.*, P.Desenho, P.Descricao from CadMaquinas_acessorios CMA INNER JOIN Projproduto P ON P.Codproduto = CMA.ID_produto where CMA.ID = " & Lista_acess.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposAcess
    Txt_ID_acess = TBAbrir!ID
    Txt_ID_produto_acess = TBAbrir!ID_Produto
    Txt_codigo_int_acess = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
    Txt_descricao_acess = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    If IsNull(TBAbrir!Operacao_prep) = False And TBAbrir!Operacao_prep <> "" Then
        If TBAbrir!Operacao_prep = 1 Then Cmb_operacao_prep_acess = "Somar" Else Cmb_operacao_prep_acess = "Subtrair"
    End If
    Txt_valor_prep_acess = IIf(IsNull(TBAbrir!Valor_prep), "", Format(TBAbrir!Valor_prep, "###,##0.00"))
    If IsNull(TBAbrir!Operacao_exec) = False And TBAbrir!Operacao_exec <> "" Then
        If TBAbrir!Operacao_exec = 1 Then Cmb_operacao_exec_acess = "Somar" Else Cmb_operacao_exec_acess = "Subtrair"
    End If
    Txt_valor_exec_acess = IIf(IsNull(TBAbrir!Valor_exec), "", Format(TBAbrir!Valor_exec, "###,##0.00"))
    Frame4.Enabled = True
    Novo_maquina3 = False
    CodigoLista3 = Lista_acess.SelectedItem.index
End If
TBAbrir.Close

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
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CadMaquinas_instrucoes where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposInstrucoes
    Txt_ID_instrucao = TBAbrir!ID
    Txt_instrucao = TBAbrir!Instrucao
    Frame2.Enabled = True
    Novo_maquina2 = False
    CodigoLista2 = Lista.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaquina_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaMaquina
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from CadMaquinas where IDMaquina = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    If Cmb_opcao_lista = "Excluir" Then
                        ProcVerificaRegistroUtilizadoSemMsg "Fases", "maquina = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                        If Permitido = False Then GoTo Proximo
                        ProcVerificaRegistroUtilizadoSemMsg "Ordemservico", "maquina = '" & .ListItems(InitFor).ListSubItems(3) & "'"
                        If Permitido = False Then GoTo Proximo
                    End If
                End If
                TBMaquinas.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaMaquina, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaquina_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaMaquina
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Mensagem = "Não é permitido excluir este posto de trabalho, pois o mesmo está sendo utilizado no módulo"
            If Cmb_opcao_lista = "Excluir" Then
                ProcVerificaRegistroUtilizado "Fases", "maquina = '" & .ListItems(InitFor).ListSubItems(3) & "'", "Engenharia/Processos"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "Ordemservico", "maquina = '" & .ListItems(InitFor).ListSubItems(3) & "'", "PCP/Gerenciamento de ordem"
                If Permitido = False Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaquina_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaMaquina.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM Cadmaquinas WHERE idmaquina = " & ListaMaquina.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    Framemaquina.Enabled = True
    Novo_maquina = False
    CodigoLista = ListaMaquina.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaTurnos()
On Error GoTo tratar_erro

USTreeView1.Clear
With USTreeView1
    .ImageList = ImageList1
    
    Set NivelP1 = .Nodes.AddNode("Dias da semana", 0, , True, , , , 2, vbBlack)
    
    Contador = 1
    Do While Contador <= 7
        Select Case Contador
            Case 1: Diasemana = "Segunda"
            Case 2: Diasemana = "Terça"
            Case 3: Diasemana = "Quarta"
            Case 4: Diasemana = "Quinta"
            Case 5: Diasemana = "Sexta"
            Case 6: Diasemana = "Sabado"
            Case 7: Diasemana = "Domingo"
        End Select
        Set NivelP2 = .Nodes.AddNode(Diasemana, Diasemana, , True, True, , , 1, vbBlack, NivelP1)
        
        Contador2 = 1
        Set TBTempo = CreateObject("adodb.recordset")
        StrSql = "Select * FROM CadmaqTurnos WHERE maquina = '" & txtmaquina & "' and Diasemana = '" & Diasemana & "'"
        'Debug.print StrSql
        
        TBTempo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBTempo.EOF = False Then
            Do While TBTempo.EOF = False
                Descricao = "Turno : " & TBTempo!Turno & " - Total turno: " & Left(TBTempo!TotalTurno, 8) & " - Status: " & IIf(TBTempo!Bloqueado = True, "Bloqueado", "Liberado")
                Set NivelP3 = .Nodes.AddNode(Descricao, Diasemana & Contador2, TBTempo!CODIGO, , , , , 3, vbBlack, NivelP2)
                Contador2 = Contador2 + 1
                TBTempo.MoveNext
            Loop
        End If
        TBTempo.Close
        
        Contador = Contador + 1
    Loop
    .ExpandAllNodes True
    .RefreshData
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mskFinal_intervalo_LostFocus()
On Error GoTo tratar_erro

If mskFinal_intervalo < mskInicio_intervalo Then
    USMsgBox ("O final não pode ser menor que o inicio do intervalo."), vbExclamation, "CAPRIND v5.0"
    mskFinal_intervalo.SetFocus
    Exit Sub
End If
ProcCalculaTemposIntervalo
ProcCalculaTempos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mskfinal_LostFocus()
On Error GoTo tratar_erro

ProcCalculaTempos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mskInicio_intervalo_LostFocus()
On Error GoTo tratar_erro

ProcCalculaTemposIntervalo
ProcCalculaTempos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mskinicio_LostFocus()
On Error GoTo tratar_erro

ProcCalculaTempos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTemposIntervalo()
On Error GoTo tratar_erro
Dim Final       As Date 'OK
Dim HoraTotal   As Date 'OK
Dim TotalTurno  As Date 'OK

If Len(mskInicio_intervalo.Value) = 8 And Len(mskFinal_intervalo.Value) = 8 Then
    If mskInicio_intervalo.Value <> "00:00:00" And mskFinal_intervalo.Value <> "00:00:00" Then
        Inicio = mskInicio_intervalo.Value
        Final = mskFinal_intervalo.Value
        If Final > Inicio Then
            HoraTotal = Final - Inicio
        Else
            Final = Final + 1
            HoraTotal = Final - Inicio
        End If
    End If
    txtIntervalo.Text = HoraTotal
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTempos()
On Error GoTo tratar_erro
Dim Final       As Date 'OK
Dim HoraTotal   As Date 'OK
Dim TotalTurno  As Date 'OK

If Len(mskinicio.Value) = 8 And Len(mskfinal.Value) = 8 Then
    If mskinicio.Value <> "00:00:00" And mskfinal.Value <> "00:00:00" Then
        Inicio = mskinicio.Value
        Final = mskfinal.Value
        If txtIntervalo.Text <> "00:00:00" And txtIntervalo.Text <> "" Then
            intervalo = txtIntervalo.Text
        End If
        If Final > Inicio Then
            HoraTotal = Final - Inicio
        Else
            Final = Final + 1
            HoraTotal = Final - Inicio
        End If
        txtTurno = HoraTotal
        If txtIntervalo.Text <> "00:00:00" And txtIntervalo.Text <> "" Then
            If intervalo < HoraTotal Then
                HoraTotal = HoraTotal - intervalo
            Else
                USMsgBox ("Intervalo maior que total do turno."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
        End If
        txttotal1.Text = HoraTotal
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtIDmaquina.Text = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0: If ListaMaquina.Visible = True Then ListaMaquina.SetFocus
    Case 1:
        Txt_setfocus.SetFocus
        If ProcVerificaProsseguir = False Then Exit Sub
        ProcCarregaTurnos
    Case 2:
        Lista.SetFocus
        If ProcVerificaProsseguir = False Then Exit Sub
        ProcCarregaListaInstrucoes
    Case 3:
        Lista_acess.SetFocus
        If ProcVerificaProsseguir = False Then Exit Sub
        ProcCarregaListaAcess
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerificaProsseguir() As Boolean
On Error GoTo tratar_erro

ProcVerificaProsseguir = True
If Novo_maquina = True Then
    USMsgBox ("Salve o posto de trabalho antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    ProcVerificaProsseguir = False
    SSTab1.Tab = 0
    Exit Function
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Txt_valor_exec_acess_Change()
On Error GoTo tratar_erro

If Txt_valor_exec_acess <> "" Then
    VerifNumero = Txt_valor_exec_acess
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_exec_acess = ""
        Txt_valor_exec_acess.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_exec_acess_LostFocus()
On Error GoTo tratar_erro

Txt_valor_exec_acess = Format(Txt_valor_exec_acess, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_prep_acess_Change()
On Error GoTo tratar_erro

If Txt_valor_prep_acess <> "" Then
    VerifNumero = Txt_valor_prep_acess
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_prep_acess = ""
        Txt_valor_prep_acess.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_prep_acess_LostFocus()
On Error GoTo tratar_erro

Txt_valor_prep_acess = Format(Txt_valor_prep_acess, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcustohora_Change()
On Error GoTo tratar_erro

If txtCustoHora.Text <> "" Then
    VerifNumero = txtCustoHora.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCustoHora.Text = ""
        txtCustoHora.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcustohora_LostFocus()
On Error GoTo tratar_erro

txtCustoHora.Text = Format(txtCustoHora.Text, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCustoHora_Setup_Change()
On Error GoTo tratar_erro

If txtCustoHora_Setup.Text <> "" Then
    VerifNumero = txtCustoHora_Setup.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtCustoHora_Setup.Text = ""
        txtCustoHora_Setup.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCustoHora_Setup_LostFocus()
On Error GoTo tratar_erro

txtCustoHora_Setup.Text = Format(txtCustoHora_Setup.Text, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDescM_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtDescM

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtHora_extra_Change()
On Error GoTo tratar_erro

If txtHora_extra.Text <> "" Then
    VerifNumero = txtHora_extra.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtHora_extra.Text = ""
        txtHora_extra.SetFocus
        Exit Sub
    End If
    valor = txtHora_extra
    If valor > 100 Then
        USMsgBox ("O percentual de hora extra não pode ser maior que 100."), vbExclamation, "CAPRIND v5.0"
        txtHora_extra = ""
        txtHora_extra.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIntervalo_LostFocus()
On Error GoTo tratar_erro

ProcCalculaTempos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaMaquina
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) posto(s) de trabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from CadMaquinas WHERE IDMaquina = " & .ListItems.Item(InitFor)
            Conexao.Execute "DELETE from CadmaqTurnos WHERE Maquina = '" & .ListItems.Item(InitFor).ListSubItems(3) & "'"
            Conexao.Execute "DELETE from CadMaquinas_instrucoes WHERE ID_Maquina = " & .ListItems.Item(InitFor)
            
            '==================================
            Modulo = "PCP/Postos de trabalho"
            Evento = "Excluir"
            ID_documento = .ListItems.Item(InitFor)
            Documento = "Código do posto de trabalho: " & .ListItems.Item(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) posto(s) de trabalho antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Posto(s) de trabalho excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    Framemaquina.Enabled = False
    ProcLimpaCampos
    ProcAtualizaListaPT (1)
    Novo_maquina = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtIDmaquina.Text = 0
Txt_date = Format(Date, "dd/mm/yy")
Txt_responsavel = pubUsuario
Txt_status = "Liberado"
Optcustos.Value = 1
Chk_liberada.Value = 0
Chk_insp_final.Value = 0
txtmaquina.Text = ""
txtGrupo.Text = ""
Txt_centro_de_custo.Text = ""
txtCustoHora.Text = ""
txtCustoHora_Setup.Text = ""
txtDescM.Text = ""
txtdados.Text = ""
CodigoLista = 0
Caption = "PCP - Postos de trabalho"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTurno()
On Error GoTo tratar_erro
    
cmbdia.ListIndex = -1
cmbturno.ListIndex = -1
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel1 = pubUsuario
txtStatus1 = "Liberado"
mskinicio.Value = "00:00:00"
Msk_margem_ap.Value = "00:00:00"
mskfinal.Value = "00:00:00"
mskInicio_intervalo.Value = "00:00:00"
mskFinal_intervalo.Value = "00:00:00"
chkIntervalo.Value = 0
txtTurno.Text = "00:00:00"
txtIntervalo.Text = "00:00:00"
txttotal1.Text = "00:00:00"
txthoras.Text = ""
txtHora_extra = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Caption = "PCP - Postos de trabalho - (Posto de trabalho : " & IIf(IsNull(TBAbrir!maquina), "", TBAbrir!maquina) & ")"
txtIDmaquina.Text = IIf(IsNull(TBAbrir!IDMaquina), 0, TBAbrir!IDMaquina)
Txt_data.Text = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
Txt_responsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
If TBAbrir!Bloqueado = True Then Txt_status = "Bloqueado" Else Txt_status = "Liberado"
If TBAbrir!custos = True Then Optcustos.Value = 1 Else Optcustos.Value = 0
If TBAbrir!Liberada = True Then Chk_liberada.Value = 1 Else Chk_liberada.Value = 0
If TBAbrir!Insp_final = True Then Chk_insp_final.Value = 1 Else Chk_insp_final.Value = 0
txtmaquina = IIf(IsNull(TBAbrir!maquina), "", TBAbrir!maquina)
txtGrupo.Text = IIf(IsNull(TBAbrir!Grupo), "", TBAbrir!Grupo)
Txt_centro_de_custo.Text = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
txtCustoHora.Text = IIf(IsNull(TBAbrir!PrecoHora), "0,00", Format(TBAbrir!PrecoHora, "###,##0.00"))
txtCustoHora_Setup.Text = IIf(IsNull(TBAbrir!PrecoHora_Setup), "0,00", Format(TBAbrir!PrecoHora_Setup, "###,##0.00"))
txtDescM = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
txtdados.Text = IIf(IsNull(TBAbrir!caracteristicas), "", TBAbrir!caracteristicas)
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtIDmaquina = "0" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CadMaquinas order by Maquina", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDMaquina = " & txtIDmaquina)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtIDmaquina.Text = TBLISTA!IDMaquina
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from cadmaquinas where idmaquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCamposInstrucoes
        ProcLimpaCamposTurno
        ProcPuxaDados
        ProcCarregaTurnos
        ProcCarregaListaInstrucoes
    Else
        USMsgBox ("Fim dos cadastros de postos de trabalho."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_maquina = False
Novo_maquina1 = False
Novo_maquina2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtIDmaquina = "0" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CadMaquinas order by Maquina", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDMaquina = " & txtIDmaquina)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtIDmaquina.Text = TBLISTA!IDMaquina
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from cadmaquinas where idmaquina = " & txtIDmaquina, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcLimpaCamposInstrucoes
        ProcLimpaCamposTurno
        ProcPuxaDados
        ProcCarregaTurnos
        ProcCarregaListaInstrucoes
    Else
        USMsgBox ("Fim dos cadastros de postos de trabalho."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_maquina = False
Novo_maquina1 = False
Novo_maquina2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=Eaa813pqSQQ")

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
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcCopiar
    Case 9: ProcStatus
    Case 10: ProcCustoEficiencia
    Case 11: ProcLiberar
    Case 12: procAtualiza
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
    Case 1: ProcNovoturno
    Case 2: ProcGravarTurno
    Case 3: ProcExcluirTurno
    Case 4: ProcImprimir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcCopiarTurno
    Case 8: ProcBloquearTurno
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoInstrucao
    Case 2: ProcGravarInstrucao
    Case 3: ProcExcluirInstrucao
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

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoAcess
    Case 2: ProcGravarAcess
    Case 3: ProcExcluirAcess
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

Private Sub USTreeView1_NodeSelected(Node As DrawSuite2022.TreeNode)
On Error GoTo tratar_erro

ProcLimpaCamposTurno
If IsNumeric(Node.ItemData) = True Then
    Framehoras.Enabled = True
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select * from CadMaqTurnos where Codigo = " & Node.ItemData, Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        txtData1 = IIf(IsNull(TBTempo!Data), "", Format(TBTempo!Data, "dd/mm/yy"))
        txtResponsavel1 = IIf(IsNull(TBTempo!Responsavel), "", TBTempo!Responsavel)
        If TBTempo!Bloqueado = True Then txtStatus1 = "Bloqueado" Else txtStatus1 = "Liberado"
        cmbdia.Text = TBTempo!Diasemana
        cmbturno.Text = TBTempo!Turno
    End If
    TBTempo.Close
    ProcCalculaTempos
    Novo_maquina1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
