VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Relatorios_Relacionamento 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Relatórios - Relacionamento de nota fiscal"
   ClientHeight    =   10035
   ClientLeft      =   1950
   ClientTop       =   1665
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   64
      Top             =   5430
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
         TabIndex        =   22
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
         Left            =   3780
         TabIndex        =   21
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   26
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Relatorios_Relacionamento.frx":0000
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
         TabIndex        =   25
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Relatorios_Relacionamento.frx":37AA
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
         TabIndex        =   23
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
         TabIndex        =   24
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Relatorios_Relacionamento.frx":72C3
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
         TabIndex        =   27
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_Relatorios_Relacionamento.frx":B3BD
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
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13050
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3090
         TabIndex        =   66
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4410
         TabIndex        =   65
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   4035
      TabIndex        =   63
      Top             =   2490
      Width           =   7275
      Begin VB.CheckBox chkSaldo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo > 0"
         Height          =   195
         Left            =   5940
         TabIndex        =   69
         Top             =   270
         Width           =   1695
      End
      Begin VB.CheckBox chkMaoObra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mão de obra"
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
         Height          =   210
         Left            =   1200
         TabIndex        =   15
         Top             =   270
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Chk_demonstracao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Demonstração"
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
         Height          =   210
         Left            =   2640
         TabIndex        =   16
         Top             =   270
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.CheckBox chkOutras 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outras operações"
         Height          =   195
         Left            =   4245
         TabIndex        =   17
         Top             =   270
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkVendas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vendas"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   270
         Value           =   1  'Checked
         Width           =   1605
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   60
      TabIndex        =   62
      Top             =   2490
      Width           =   3945
      Begin VB.CheckBox Chk_integral 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Integral"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2910
         TabIndex        =   13
         Top             =   270
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.CheckBox Chk_parcial 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Parcial"
         Height          =   195
         Left            =   1650
         TabIndex        =   12
         Top             =   270
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox Chk_relacionar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Relacionar"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Value           =   1  'Checked
         Width           =   1065
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   60
      TabIndex        =   58
      Top             =   1620
      Width           =   1845
      Begin VB.OptionButton optEntrada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   900
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
      Begin VB.OptionButton optSaida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saída"
         DisabledPicture =   "frmFaturamento_Relatorios_Relacionamento.frx":EC8A
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   1920
      TabIndex        =   53
      Top             =   1620
      Width           =   13275
      Begin VB.ComboBox cmbfiltrarpor 
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
         ItemData        =   "frmFaturamento_Relatorios_Relacionamento.frx":258BCC
         Left            =   180
         List            =   "frmFaturamento_Relatorios_Relacionamento.frx":258BE5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3060
         TabIndex        =   54
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
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
            Left            =   2760
            TabIndex        =   9
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
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
            TabIndex        =   7
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
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
            Left            =   1470
            TabIndex        =   8
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
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
            Left            =   3930
            TabIndex        =   10
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.TextBox txtTexto 
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
         Left            =   7950
         TabIndex        =   4
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   5145
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
         ItemData        =   "frmFaturamento_Relatorios_Relacionamento.frx":258C44
         Left            =   7950
         List            =   "frmFaturamento_Relatorios_Relacionamento.frx":258C46
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Visible         =   0   'False
         Width           =   5145
      End
      Begin VB.Label Label5 
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
         Left            =   1567
         TabIndex        =   56
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9787
         TabIndex        =   55
         Top             =   180
         Width           =   1470
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00E0E0E0&
      Height          =   630
      Left            =   60
      TabIndex        =   51
      Top             =   990
      Width           =   15195
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
         ItemData        =   "frmFaturamento_Relatorios_Relacionamento.frx":258C48
         Left            =   1170
         List            =   "frmFaturamento_Relatorios_Relacionamento.frx":258C4A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   13845
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
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
         Left            =   240
         TabIndex        =   52
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por período"
      Height          =   195
      Left            =   11475
      TabIndex        =   18
      Top             =   2490
      Width           =   1245
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   38
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
      Height          =   3645
      Left            =   60
      TabIndex        =   37
      Top             =   6060
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14215660
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
      TabPicture(0)   =   "frmFaturamento_Relatorios_Relacionamento.frx":258C4C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListView2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notas fiscais relacionadas"
      TabPicture(1)   =   "frmFaturamento_Relatorios_Relacionamento.frx":258C68
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtidproduto"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtidproduto 
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
         Left            =   2580
         TabIndex        =   49
         ToolTipText     =   "Nº da nota."
         Top             =   1620
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   45
         TabIndex        =   45
         Top             =   330
         Width           =   15105
         Begin VB.TextBox txtQtd 
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
            Left            =   13590
            TabIndex        =   32
            ToolTipText     =   "Quantidade."
            Top             =   375
            Width           =   1305
         End
         Begin VB.TextBox txtCodref 
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
            Left            =   2130
            TabIndex        =   30
            ToolTipText     =   "Código de referência."
            Top             =   375
            Width           =   2685
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
            Height          =   315
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   375
            Width           =   8745
         End
         Begin VB.TextBox txtCodinterno 
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
            TabIndex        =   29
            ToolTipText     =   "Código interno."
            Top             =   375
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   13815
            TabIndex        =   50
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. interno"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   697
            TabIndex        =   48
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   8857
            TabIndex        =   47
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   2910
            TabIndex        =   46
            Top             =   180
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   45
         TabIndex        =   39
         Top             =   2775
         Width           =   15105
         Begin VB.TextBox txtQtde1 
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
            Height          =   315
            Left            =   4500
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "0,000"
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtQtdeRel 
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
            Height          =   315
            Left            =   6615
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade relacionada."
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtSaldo 
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
            Height          =   315
            Left            =   8760
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Saldo"
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. entrada"
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
            Left            =   4725
            TabIndex        =   44
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   6345
            TabIndex        =   43
            Top             =   450
            Width           =   75
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. relacionada"
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
            Left            =   6690
            TabIndex        =   42
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   8430
            TabIndex        =   41
            Top             =   450
            Width           =   135
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   9345
            TabIndex        =   40
            Top             =   180
            Width           =   465
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   -74955
         TabIndex        =   28
         Top             =   345
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483641
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
            Object.Width           =   0
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
            Text            =   "Cód. de ref."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   15020
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. relac."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1575
         Left            =   45
         TabIndex        =   33
         Top             =   1200
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483641
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Nota fiscal"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Destinatário/Emitente"
            Object.Width           =   17313
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. relac."
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2250
      Left            =   60
      TabIndex        =   6
      Top             =   3180
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   3969
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483641
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   19465
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   57
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   6
      GradientColor1  =   16777215
      GradientColor2  =   14737632
      GradientColorDown1=   10802943
      GradientColorDown2=   7979263
      GradientColorDownRight1=   10802943
      GradientColorDownRight2=   7979263
      GradientColorOver1=   14417407
      GradientColorOver2=   12317439
      GradientColorOverRight1=   14417407
      GradientColorOverRight2=   12317439
      IsStrech        =   -1  'True
      RightColor1     =   14737632
      RightColor2     =   16777215
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   51
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
      ButtonLeft3     =   93
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
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
      ButtonState6    =   5
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11640
         Top             =   195
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Relatorios_Relacionamento.frx":258C84
         Count           =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   11340
      TabIndex        =   59
      Top             =   2490
      Width           =   3855
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   2370
         TabIndex        =   20
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   270
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
         Format          =   197328897
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   540
         TabIndex        =   19
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   270
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
         Format          =   197328897
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1935
         TabIndex        =   61
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   60
         Top             =   270
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmFaturamento_Relatorios_Relacionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Faturamento_Relatorios_Relacionamento    As String 'OK
Dim FiltroRel_Faturamento_Relatorios_Relacionamento    As String 'OK
Dim TBLISTA_Faturamento_Relatorios_Relacionamento      As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=o9mVNykTaq0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=10&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ID_nota, Int_codigo, int_Cod_Produto, N_Referencia, Txt_descricao, int_Qtd from tbl_Detalhes_Nota where id_nota = " & ListView1.SelectedItem & " order by Int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Int_codigo = " & txtidproduto)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        ProcLimparCamposProdutos
        txtidproduto = TBLISTA!Int_codigo
        txtCodinterno = IIf(IsNull(TBLISTA!int_Cod_Produto), "", TBLISTA!int_Cod_Produto)
        txtCodref = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
        txtdescricao = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
        txtQTD = IIf(IsNull(TBLISTA!int_Qtd), "", Format(TBLISTA!int_Qtd, "###,##0.0000"))
        If ProcVerifNFComplementar(TBLISTA!ID_nota) = True Then ProcCarregaListaRelacionada True Else ProcCarregaListaRelacionada False
    Else
        USMsgBox ("Fim dos cadastros de produtos."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "Faturamento_relacionamento.rpt"
ProcImprimirRel FiltroRel_Faturamento_Relatorios_Relacionamento, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ID_nota, Int_codigo, int_Cod_Produto, N_Referencia, Txt_descricao, int_Qtd from tbl_Detalhes_Nota where id_nota = " & ListView1.SelectedItem & " order by Int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Int_codigo = " & txtidproduto)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimparCamposProdutos
        txtidproduto = TBLISTA!Int_codigo
        txtCodinterno = IIf(IsNull(TBLISTA!int_Cod_Produto), "", TBLISTA!int_Cod_Produto)
        txtCodref = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
        txtdescricao = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
        txtQTD = IIf(IsNull(TBLISTA!int_Qtd), "", Format(TBLISTA!int_Qtd, "###,##0.0000"))
        If ProcVerifNFComplementar(TBLISTA!ID_nota) = True Then ProcCarregaListaRelacionada True Else ProcCarregaListaRelacionada False
    Else
        USMsgBox ("Fim dos cadastros de produtos."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_demonstracao_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_integral_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_parcial_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkMaoObra_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkOutras_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVendas_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear
If cmbfiltrarpor = "Família produto" Or cmbfiltrarpor = "Família serviço" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
Else
    txtTexto.Visible = True
    cmbfamilia.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage <> 2 Then
    If TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount - 1)
    Else
        TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage)
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
    TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = 1
ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage <> -3 Then
    If TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
ProcExibePagina (TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and (compras = 'True' or vendas = 'True')", True
cmbfiltrarpor = "Nota fiscal"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date
ProcCorrigeForm (False)
Status_nota = 1
SSTab1.Tab = 0

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaNF()
On Error GoTo tratar_erro

If StrSql_Faturamento_Relatorios_Relacionamento = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
ListView2.ListItems.Clear
Set TBLISTA_Faturamento_Relatorios_Relacionamento = CreateObject("adodb.recordset")
TBLISTA_Faturamento_Relatorios_Relacionamento.Open StrSql_Faturamento_Relatorios_Relacionamento, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Faturamento_Relatorios_Relacionamento.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Faturamento_Relatorios_Relacionamento.RecordCount - IIf(Pagina > 1, (TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize * (Pagina - 1)), 0), TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Faturamento_Relatorios_Relacionamento.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLISTA_Faturamento_Relatorios_Relacionamento!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!dt_DataEmissao), "", Format(TBLISTA_Faturamento_Relatorios_Relacionamento!dt_DataEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!int_NotaFiscal), "", TBLISTA_Faturamento_Relatorios_Relacionamento!int_NotaFiscal)
        If IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF) = False Then
            If TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
            If TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(3) = TipoNF2
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!txt_Razao_Nome), "", Trim(TBLISTA_Faturamento_Relatorios_Relacionamento!txt_Razao_Nome))
    End With
    TBLISTA_Faturamento_Relatorios_Relacionamento.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Faturamento_Relatorios_Relacionamento.RecordCount
If TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
ElseIf TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount & " de: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage - 1 & " de: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCamposProdutos()
On Error GoTo tratar_erro

txtidproduto = ""
txtCodinterno = ""
txtCodref = ""
txtdescricao = ""
txtQTD = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaProdutos(NFcomplementar As Boolean)
On Error GoTo tratar_erro

ListView2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Int_codigo, int_Cod_Produto, N_Referencia, Txt_descricao, int_Qtd, Saldo from tbl_Detalhes_Nota where id_nota = " & ListView1.SelectedItem & " order by int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListView2.ListItems
            .Add , , TBLISTA!Int_codigo
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!int_Cod_Produto), "", TBLISTA!int_Cod_Produto)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Txt_descricao), "", Trim(TBLISTA!Txt_descricao))
            If NFcomplementar = False Then
                Qtde = IIf(IsNull(TBLISTA!int_Qtd), 0, TBLISTA!int_Qtd)
                .Item(.Count).SubItems(4) = Format(Qtde, "###,##0.0000")
                quantidade = IIf(IsNull(TBLISTA!Saldo), 0, TBLISTA!Saldo)
                .Item(.Count).SubItems(5) = Format(Qtde - quantidade, "###,##0.0000")
                .Item(.Count).SubItems(6) = Format(quantidade, "###,##0.0000")
            End If
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close
ProcLimparCamposProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaRelacionada(NFcomplementar As Boolean)
On Error GoTo tratar_erro

ListView3.ListItems.Clear
Qtde = IIf(txtQTD = "", 0, txtQTD)
quantidade = 0

If NFcomplementar = True Then
    TextoFiltro = "ID_nota = " & ListView1.SelectedItem & " or ID_nota_relacionada = " & ListView1.SelectedItem
Else
    TextoFiltro = "ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & txtidproduto & " or ID_nota_relacionada = " & ListView1.SelectedItem & " and ID_produto_relacionada = " & txtidproduto
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBAbrir.EOF = False
        With ListView3.ListItems
            .Add , , TBAbrir!ID
            
            With frmFaturamento_Prod_Serv
                If TBAbrir!ID_nota = ListView1.SelectedItem Then
                    If NFcomplementar = True Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota_relacionada Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!id_produto_relacionada
                Else
                    If NFcomplementar = True Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!ID_Produto
                End If
            End With
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome, NFP.dbl_ValorUnitario, NFP.Unidade_com from tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!dt_DataEmissao), "", (Format(TBFI!dt_DataEmissao, "dd/mm/yy")))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBFI!int_NotaFiscal), "", TBFI!int_NotaFiscal)
                If IsNull(TBFI!TipoNF) = False Then
                    If TBFI!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
                    If TBFI!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
                    If TBFI!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
                End If
                .Item(.Count).SubItems(3) = TipoNF2
                .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!txt_Razao_Nome), "", TBFI!txt_Razao_Nome)
                
                If NFcomplementar = False Then
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Qtde), 0, Format(TBAbrir!Qtde, "###,##0.0000"))
                    '.Item(.Count).SubItems(6) = IIf(IsNull(TBFI!dbl_ValorUnitario), 0, Format(TBFI!dbl_ValorUnitario, "###,##0.00000"))
                    '.Item(.Count).SubItems(7) = IIf(IsNull(TBFI!Unidade_com), 0, TBFI!Unidade_com)
                End If
            End If
            TBFI.Close
            
            quantidade = quantidade + TBAbrir!Qtde
        End With
        TBAbrir.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBAbrir.Close
txtQtde1 = Format(Qtde, "###,##0.0000")
txtQtdeRel = Format(quantidade, "###,##0.0000")
txtSaldo = Format(Qtde - quantidade, "###,##0.0000")
    
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

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView1
    If .ListItems.Count = 0 Then Exit Sub
    SSTab1.Tab = 0
    If ProcVerifNFComplementar(.SelectedItem) = True Then
        With ListView2
            .ColumnHeaders(4).Width = 12235
            .ColumnHeaders(5).Width = 0
            .ColumnHeaders(6).Width = 0
            .ColumnHeaders(7).Width = 0
        End With
        ProcCarregaListaProdutos True
    Else
        With ListView2
            .ColumnHeaders(4).Width = 8635
            .ColumnHeaders(5).Width = 1200
            .ColumnHeaders(6).Width = 1200
            .ColumnHeaders(7).Width = 1200
        End With
        ProcCarregaListaProdutos False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView2, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView2
    If .ListItems.Count = 0 Then Exit Sub
    ProcLimparCamposProdutos
    txtidproduto = .SelectedItem
    txtCodinterno = .SelectedItem.ListSubItems(1)
    txtCodref = .SelectedItem.ListSubItems(2)
    txtdescricao = .SelectedItem.ListSubItems(3)
    txtQTD = .SelectedItem.ListSubItems(4)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView3, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optEntrada_Click()
On Error GoTo tratar_erro
    
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ProcCorrigeForm (True)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeForm(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Nota fiscal"
    
    If Entrada = True Then
        .AddItem "Emitente"
        .Text = "Emitente"
    Else
        .AddItem "Destinatário"
        .Text = "Destinatário"
    End If
End With

If Entrada = True Then
    With Label1
        .Caption = "Qtde. entrada"
        .Left = txtQtde1.Left + (txtQtde1.Width / 6)
    End With
    txtQtde1.ToolTipText = "Quantidade de entrada"
    ListView1.ColumnHeaders(5).Text = "Emitente"
    With ListView2
        .ColumnHeaders(5).Text = "Qtde. entr."
        .ColumnHeaders(6).Text = "Qtde. saída"
    End With
Else
    With Label1
        .Caption = "Qtde. saída"
        .Left = txtQtde1.Left + (txtQtde1.Width / 6)
    End With
    txtQtde1.ToolTipText = "Quantidade de saída"
    ListView1.ColumnHeaders(5).Text = "Destinatário"
    With ListView2
        .ColumnHeaders(5).Text = "Qtde. saída"
        .ColumnHeaders(6).Text = "Qtde. entr."
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear
If optPeriodo.Value = 1 Then
    Frame7.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame7.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_relacionar_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSaida_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear
ProcCorrigeForm (False)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        If ListView1.Visible = True Then ListView1.SetFocus
    Case 1:
        ListView3.SetFocus
        ListView3.ListItems.Clear
        If ListView2.ListItems.Count = 0 Then
            SSTab1.Tab = 0
            USMsgBox ("Informe a nota fiscal na lista, antes de verificar o relacionamento."), vbExclamation, "CAPRIND v5.0"
            ListView1.SetFocus
            Exit Sub
        End If
        If ProcVerifNFComplementar(ListView1.SelectedItem) = True Then
            Frame5.Visible = False
            With ListView3
                .ColumnHeaders(5).Width = 11135
                .ColumnHeaders(6).Width = 0
                .Top = USToolBar2.Top + USToolBar2.Height
                .Height = PBLista.Top - .Top
            End With
            Frame3.Visible = False
            ProcCarregaListaRelacionada True
        Else
            Frame5.Visible = True
            With ListView3
                .ColumnHeaders(5).Width = 9935
                .ColumnHeaders(6).Width = 1200
                .Top = Frame5.Top + Frame5.Height
                .Height = Frame3.Top - .Top
            End With
            Frame3.Visible = True
            If txtCodinterno = "" Then
                SSTab1.Tab = 0
                USMsgBox ("Informe o produto na lista, antes de verificar o relacionamento."), vbExclamation, "CAPRIND v5.0"
                ListView2.SetFocus
                Exit Sub
            End If
            ProcCarregaListaRelacionada False
        End If
End Select

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

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"
If chkVendas.Value = 0 And chkMaoObra.Value = 0 And Chk_demonstracao.Value = 0 And chkOutras.Value = 0 Then
    NomeCampo = "a operação"
    ProcVerificaAcao
    Exit Sub
End If
If Chk_relacionar.Value = 0 And Chk_parcial.Value = 0 And Chk_integral.Value = 0 Then
    NomeCampo = "uma das opções de relacionamento"
    ProcVerificaAcao
    Exit Sub
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

If optSaida.Value = True Then
    Sit_Nota = 1
    Aplicacao_nota = ""
    Aplicacao_notaRel = ""
Else
    Sit_Nota = 2
    Aplicacao_nota = " and NF.Aplicacao = 'T'"
    Aplicacao_notaRel = " and {tbl_Dados_Nota_Fiscal.Aplicacao} = 'T'"
End If
If chkVendas.Value = 1 Then
    CFOPVendas = "CFOP.Vendas = 'True'"
    CFOPVendasRel = "{tbl_NaturezaOperacao.Vendas} = True"
Else
    CFOPVendas = ""
    CFOPVendasRel = ""
End If
If chkMaoObra.Value = 1 Then
    If CFOPVendas = "" Then
        CFOPMO = "CFOP.MaoObra = 'True'"
        CFOPMORel = "{tbl_NaturezaOperacao.MaoObra} = True"
    Else
        CFOPMO = " or CFOP.MaoObra = 'True'"
        CFOPMORel = " or {tbl_NaturezaOperacao.MaoObra} = True"
    End If
Else
    CFOPMO = ""
    CFOPMORel = ""
End If
If Chk_demonstracao.Value = 1 Then
    If CFOPVendas = "" And CFOPMO = "" Then
        CFOPDEM = "CFOP.Demonstracao = 'True'"
        CFOPDEMRel = "{tbl_NaturezaOperacao.Demonstracao} = True"
    Else
        CFOPDEM = " or CFOP.Demonstracao = 'True'"
        CFOPDEMRel = " or {tbl_NaturezaOperacao.Demonstracao} = True"
    End If
Else
    CFOPDEM = ""
    CFOPDEMRel = ""
End If
If chkOutras.Value = 1 Then
    If CFOPVendas = "" And CFOPMO = "" And CFOPDEM = "" Then
        CFOPOutros = "CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False'"
        CFOPOutrosRel = "{tbl_NaturezaOperacao.Vendas} = False and {tbl_NaturezaOperacao.MaoObra} = False and {tbl_NaturezaOperacao.Demonstracao} = False"
    Else
        CFOPOutros = " or CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False'"
        CFOPOutrosRel = " or {tbl_NaturezaOperacao.Vendas} = False and {tbl_NaturezaOperacao.MaoObra} = False and {tbl_NaturezaOperacao.Demonstracao} = False"
    End If
Else
    CFOPOutros = ""
    CFOPOutrosRel = ""
End If

If chkVendas.Value = 1 And chkMaoObra.Value = 1 And Chk_demonstracao.Value = 1 And chkOutras.Value = 1 Then
    CFOP = "CFOP.id_CFOP <> 'Null'"
    CFOPRel = "{tbl_NaturezaOperacao.id_CFOP} <> 'Null'"
Else
    CFOP = "(" & CFOPVendas & CFOPMO & CFOPDEM & CFOPOutros & ")"
    CFOPRel = "(" & CFOPVendasRel & CFOPMORel & CFOPDEMRel & CFOPOutrosRel & ")"
End If

If Chk_relacionar.Value = 1 And Chk_parcial.Value = 1 And Chk_integral.Value = 1 Then

 If chkSaldo.Value = False Then
    Relacionamento = "(NFP.Saldo = NFP.int_Qtd or NFP.Saldo < NFP.int_Qtd or NFP.Saldo = 0)"
    Relacionamento_Rel = "({tbl_Detalhes_Nota.Saldo} = {tbl_Detalhes_Nota.int_Qtd} or {tbl_Detalhes_Nota.Saldo} < {tbl_Detalhes_Nota.int_Qtd} or {tbl_Detalhes_Nota.Saldo} = 0)"
 Else
    Relacionamento = "(NFP.Saldo > 0)"
    Relacionamento_Rel = "({tbl_Detalhes_Nota.Saldo} > 0)"
 End If
 

ElseIf Chk_relacionar.Value = 1 And Chk_parcial.Value = 1 And Chk_integral.Value = 0 Then
        Relacionamento = "(NFP.Saldo = NFP.int_Qtd or NFP.Saldo > 0 and NFP.Saldo < NFP.int_Qtd)"
        Relacionamento_Rel = "({tbl_Detalhes_Nota.Saldo} = {tbl_Detalhes_Nota.int_Qtd} or {tbl_Detalhes_Nota.Saldo} > 0 and {tbl_Detalhes_Nota.Saldo} < {tbl_Detalhes_Nota.int_Qtd})"
    ElseIf Chk_relacionar.Value = 1 And Chk_parcial.Value = 0 And Chk_integral.Value = 1 Then
            Relacionamento = "(NFP.Saldo = NFP.int_Qtd or NFP.Saldo = 0)"
            Relacionamento_Rel = "({tbl_Detalhes_Nota.Saldo} = {tbl_Detalhes_Nota.int_Qtd} or {tbl_Detalhes_Nota.Saldo} = 0)"
    ElseIf Chk_relacionar.Value = 0 And Chk_parcial.Value = 1 And Chk_integral.Value = 1 Then
                Relacionamento = "(NFP.Saldo < NFP.int_Qtd or NFP.Saldo = 0)"
                Relacionamento_Rel = "({tbl_Detalhes_Nota.Saldo} < {tbl_Detalhes_Nota.int_Qtd} or {tbl_Detalhes_Nota.Saldo} = 0)"
        ElseIf Chk_relacionar.Value = 1 Then
                Relacionamento = "NFP.Saldo = NFP.int_Qtd"
                Relacionamento_Rel = "{tbl_Detalhes_Nota.Saldo} = {tbl_Detalhes_Nota.int_Qtd}"
            ElseIf Chk_parcial.Value = 1 Then
                    Relacionamento = "NFP.Saldo > 0 and NFP.Saldo < NFP.int_Qtd"
                    Relacionamento_Rel = "{tbl_Detalhes_Nota.Saldo} > 0 and {tbl_Detalhes_Nota.Saldo} < {tbl_Detalhes_Nota.int_Qtd}"
                Else
                    Relacionamento = "NFP.Saldo = 0"
                    Relacionamento_Rel = "{tbl_Detalhes_Nota.Saldo} = 0"
End If

If optPeriodo.Value = 1 Then
    If optEntrada.Value = True Then DataFiltroTexto = "dt_Saida_Entrada" Else DataFiltroTexto = "dt_DataEmissao"
    
    DataFiltro = "(NF." & DataFiltroTexto & ") Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = "{tbl_Dados_Nota_Fiscal." & DataFiltroTexto & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Dados_Nota_Fiscal." & DataFiltroTexto & "} <= Date(" & _
                            Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
Else
    DataFiltro = "NFP.int_NotaFiscal <> 'Null'"
    DataFiltroRel = "{tbl_Dados_Nota_Fiscal.int_NotaFiscal} <> 'Null'"
End If


CamposFiltro = "NF.ID, NF.dt_dataemissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome"
INNERJOINTEXTO = "(tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NFP.ID_nota = NF.ID) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_cfop"
TextoFiltroPadrao = "NF.int_status = 1 and NF.int_TipoNota = " & Sit_Nota & Aplicacao_nota & " and " & CFOP & " and " & DataFiltro & " and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Relacionamento & " group by " & CamposFiltro & " order by NF.dt_dataemissao, NF.int_NotaFiscal"
TextoFiltroPadraoRel = "{tbl_Dados_Nota_Fiscal.int_status} = 1 and {tbl_Dados_Nota_Fiscal.int_TipoNota} = " & Sit_Nota & Aplicacao_notaRel & " and " & CFOPRel & " and " & DataFiltroRel & " and {tbl_Dados_Nota_Fiscal.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & Relacionamento_Rel

If txtTexto <> "" Or cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Then
        StrSql_Faturamento_Relatorios_Relacionamento = "Select " & CamposFiltro & " FROM " & INNERJOINTEXTO & " where NFP.Familia = '" & cmbfamilia & "' and " & TextoFilroPadrao
        FiltroRel_Faturamento_Relatorios_Relacionamento = "{tbl_detalhes_nota.Familia} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
    Else
        Select Case cmbfiltrarpor
            Case "Nota fiscal":
                TextoFiltro = "NF.int_NotaFiscal"
                If txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
                TextoFiltro_Rel = "{tbl_Dados_Nota_Fiscal.int_NotaFiscal}"
            Case "Destinatário":
                TextoFiltro = "NF.txt_Razao_Nome"
                TextoFiltro_Rel = "{tbl_Dados_Nota_Fiscal.txt_Razao_Nome}"
            Case "Emitente":
                TextoFiltro = "NF.txt_Razao_Nome"
                TextoFiltro_Rel = "{tbl_Dados_Nota_Fiscal.txt_Razao_Nome}"
            Case "CFOP":
                TextoFiltro = "CFOP.id_CFOP"
                TextoFiltro_Rel = "{tbl_NaturezaOperacao.id_CFOP}"
            Case "Código interno":
                TextoFiltro = "NFP.int_Cod_Produto"
                TextoFiltro_Rel = "{tbl_detalhes_nota.int_Cod_Produto}"
            Case "Código de referência":
                TextoFiltro = "NFP.N_Referencia"
                TextoFiltro_Rel = "{tbl_detalhes_nota.N_Referencia}"
            Case "Descrição":
                TextoFiltro = "NFP.txt_Descricao"
                TextoFiltro_Rel = "{tbl_detalhes_nota.txt_Descricao}"
        End Select
        StrSql_Faturamento_Relatorios_Relacionamento = "Select " & CamposFiltro & " FROM " & INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
        FiltroRel_Faturamento_Relatorios_Relacionamento = TextoFiltro_Rel & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
    End If
Else
    StrSql_Faturamento_Relatorios_Relacionamento = "Select " & CamposFiltro & " FROM " & INNERJOINTEXTO & " where " & TextoFiltroPadrao
    FiltroRel_Faturamento_Relatorios_Relacionamento = TextoFiltroPadraoRel
End If
'Debug.print StrSql_Faturamento_Relatorios_Relacionamento

ProcCarregaListaNF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear
If txtTexto.Text <> "" And cmbfiltrarpor = "Nota fiscal" Then
    cmbfamilia.ListIndex = -1
    VerifNumero = txtTexto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto.Text = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifNFComplementar(ID_nota As Long) As Boolean
On Error GoTo tratar_erro

ProcVerifNFComplementar = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & ID_nota & " and Finalidade_emissao = 2", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then ProcVerifNFComplementar = True
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcCarregaComboFiltrarPor(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    If Entrada = True Then .AddItem "Emitente" Else .AddItem "Destinatário"
    .AddItem "Família"
    .AddItem "Nota fiscal"
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
