VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_Natureza_OP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Fiscal - Natureza da operação"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   525
   ClientWidth     =   15330
   ControlBox      =   0   'False
   Icon            =   "frm_Natureza_OP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15330
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
      FormWidthDT     =   15450
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15330
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   35
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
      BarColor1       =   10114859
      BarColor2       =   10114859
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Dados principais"
      TabPicture(0)   =   "frm_Natureza_OP.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lst_NatOp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USToolBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtID"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Texto dados adicionais da CFOP por cliente"
      TabPicture(1)   =   "frm_Natureza_OP.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtID_CFOP_Cliente"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "ListaCliente"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outras opções"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   11190
         TabIndex        =   122
         Top             =   1320
         Width           =   4095
         Begin VB.CheckBox chkSuframa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Desconto SUFRAMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2220
            TabIndex        =   128
            Top             =   495
            Width           =   1680
         End
         Begin VB.CheckBox Chk_soma_retorno_total 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Soma retorno nos totais ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   120
            TabIndex        =   127
            Top             =   255
            Width           =   2070
         End
         Begin VB.CheckBox Chk_MPA 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Matéria-prima aplicada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2220
            TabIndex        =   126
            Top             =   240
            Width           =   1680
         End
         Begin VB.CheckBox chkReducao_BC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem redução (BC ICMS) ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   125
            Top             =   765
            Width           =   2070
         End
         Begin VB.CheckBox Chk_somar_IPI_BC_ICMSST 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Soma IPI (BC ICMS ST) ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   124
            Top             =   495
            Width           =   1980
         End
         Begin VB.CheckBox chkCreditaCentroCusto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Credita centro de custo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2220
            TabIndex        =   123
            Top             =   750
            Width           =   1800
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opções para cálculo dos impostos"
         Enabled         =   0   'False
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
         Height          =   1065
         Left            =   7620
         TabIndex        =   71
         Top             =   1320
         Width           =   3570
         Begin VB.CheckBox chk_PIS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem PIS ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   1710
            TabIndex        =   12
            ToolTipText     =   "Se selecionado calcula o pis."
            Top             =   487
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.CheckBox chk_COFINS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem COFINS ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   1710
            TabIndex        =   13
            ToolTipText     =   "Se selecionado calcula o Cofins."
            Top             =   735
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.CheckBox Chk_retem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem impostos ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Se selecionado calcula o icms, ipi pis e cofins."
            Top             =   240
            Width           =   1500
         End
         Begin VB.CheckBox chk_ICMS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem ICMS ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Se selecionado calcula o icms."
            Top             =   487
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox chk_IPI 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tem IPI ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Se selecionado calcula o ipi."
            Top             =   735
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox chk_Somar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Soma IPI (BC ICMS) ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   1710
            TabIndex        =   14
            ToolTipText     =   "Se selecionado soma o ipi na base de cálculo do icms."
            Top             =   240
            Visible         =   0   'False
            Width           =   1770
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CST's ICMS, IPI, PIS Cofins válidas"
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
         Height          =   2730
         Left            =   12240
         TabIndex        =   117
         Top             =   2370
         Width           =   3045
         Begin DrawSuite2022.USButton btnCST 
            Height          =   435
            Left            =   180
            TabIndex        =   119
            ToolTipText     =   "Cadastrar CST da CFOP"
            Top             =   2220
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   767
            DibPicture      =   "frm_Natureza_OP.frx":0044
            Caption         =   "Cadastro de CST´s"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin MSComctlLib.ListView Lista 
            Height          =   1800
            Left            =   180
            TabIndex        =   118
            ToolTipText     =   "Lista de CST's permitidas pela CFOP"
            Top             =   330
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   3175
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483641
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "ICMS"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "IPI"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "PIS"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "COFINS"
               Object.Width           =   1411
            EndProperty
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo CFOP*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   60
         TabIndex        =   110
         Top             =   1320
         Width           =   1080
         Begin VB.OptionButton OptSaida 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saida"
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
            Left            =   60
            TabIndex        =   112
            Top             =   630
            Width           =   975
         End
         Begin VB.OptionButton optEntrada 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Entrada"
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
            Left            =   60
            TabIndex        =   111
            Top             =   390
            Width           =   975
         End
      End
      Begin VB.TextBox txtID_CFOP_Cliente 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73260
         TabIndex        =   104
         Text            =   "0"
         ToolTipText     =   "Natureza da operação"
         Top             =   4500
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtID 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3600
         TabIndex        =   101
         Text            =   "0"
         ToolTipText     =   "Natureza da operação"
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   90
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
            ItemData        =   "frm_Natureza_OP.frx":A167
            Left            =   6990
            List            =   "frm_Natureza_OP.frx":A171
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   107
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
            Left            =   2730
            TabIndex        =   92
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
            TabIndex        =   91
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   93
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_Natureza_OP.frx":A189
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
            TabIndex        =   94
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_Natureza_OP.frx":D92D
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
            TabIndex        =   95
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
            TabIndex        =   96
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_Natureza_OP.frx":11436
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
            TabIndex        =   97
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frm_Natureza_OP.frx":15525
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
         Begin VB.Label Label13 
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
            TabIndex        =   109
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
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
            Index           =   29
            Left            =   5610
            TabIndex        =   108
            Top             =   233
            Width           =   1260
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
            TabIndex        =   100
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
            TabIndex        =   99
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
            TabIndex        =   98
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Utilize as variáveis abaixo para compor o texto padrão de aproveitamento de crédito do simples nacional"
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
         Height          =   615
         Left            =   75
         TabIndex        =   89
         Top             =   5100
         Width           =   15195
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11310
            TabIndex        =   116
            Text            =   "@NfVlrAproxTrib valor aproximado dos tributos"
            Top             =   270
            Width           =   3585
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   390
            TabIndex        =   115
            Text            =   "@NfVlrICMSSN Valor do ICMS SN"
            Top             =   270
            Width           =   2475
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4020
            TabIndex        =   114
            Text            =   "@NfAliqICMSSN Alíquota do ICMS SN"
            Top             =   270
            Width           =   2835
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8010
            TabIndex        =   113
            Text            =   "@NfVlrTotal  Valor total da nota  "
            Top             =   270
            Width           =   2595
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   2190
         Left            =   -74925
         TabIndex        =   83
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton cmdCliente 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14670
            Picture         =   "frm_Natureza_OP.frx":18DB1
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Localizar cliente."
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox txtID_cliente 
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
            Left            =   4980
            TabIndex        =   28
            ToolTipText     =   "Código do cliente."
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox txtRazao 
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
            Left            =   5850
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Razão social."
            Top             =   375
            Width           =   8805
         End
         Begin VB.TextBox txtResponsavel_cliente 
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3645
         End
         Begin VB.TextBox txtData_cliente 
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
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   1125
         End
         Begin VB.TextBox txtCorpoNota_cliente 
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
            Height          =   1125
            Left            =   7620
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            ToolTipText     =   "Texto padrão para corpo da nota."
            Top             =   945
            Width           =   7365
         End
         Begin VB.TextBox txtDadosAdicionais_cliente 
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
            Height          =   1125
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            ToolTipText     =   "Texto padrão para dados adicionais da nota."
            Top             =   950
            Width           =   7395
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Razão social*"
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
            Left            =   9765
            TabIndex        =   103
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
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
            Left            =   5325
            TabIndex        =   102
            Top             =   180
            Width           =   165
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
            Index           =   2
            Left            =   2685
            TabIndex        =   87
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   570
            TabIndex        =   86
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Texto corpo da nota (Reservado ao Fisco)"
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
            Left            =   9780
            TabIndex        =   85
            Top             =   750
            Width           =   3045
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Texto dados adicionais (Utilizado na NFe)"
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
            Left            =   2400
            TabIndex        =   84
            Top             =   750
            Width           =   2955
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aplicação*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1155
         TabIndex        =   74
         Top             =   1320
         Width           =   1410
         Begin VB.CheckBox chk_Proprio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nota própria"
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
            Left            =   60
            TabIndex        =   0
            Top             =   375
            Width           =   1245
         End
         Begin VB.CheckBox chk_Terceiros 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nota terceiro"
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
            Left            =   60
            TabIndex        =   1
            Top             =   630
            Width           =   1275
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Destino*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2580
         TabIndex        =   73
         Top             =   1320
         Width           =   1725
         Begin VB.OptionButton optFE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fora do estado"
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
            Left            =   60
            TabIndex        =   3
            Top             =   630
            Width           =   1605
         End
         Begin VB.OptionButton optDE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dentro do estado"
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
            Left            =   60
            TabIndex        =   2
            Top             =   375
            Width           =   1635
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo da operação fiscal (CFOP)"
         Enabled         =   0   'False
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
         Height          =   1065
         Left            =   4320
         TabIndex        =   72
         Top             =   1320
         Width           =   3285
         Begin VB.CheckBox Chk_remessa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(5) Remessa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1830
            TabIndex        =   8
            Top             =   540
            Width           =   1410
         End
         Begin VB.CheckBox Chk_retorno 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(6) Retorno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1830
            TabIndex        =   9
            Top             =   780
            Width           =   1410
         End
         Begin VB.CheckBox Chk_demonstracao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(3) Demonstração"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   6
            Top             =   780
            Width           =   1860
         End
         Begin VB.CheckBox chkMaoObra 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(2) Industrialização"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   5
            Top             =   525
            Width           =   1710
         End
         Begin VB.CheckBox Chkvendas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(1) Vendas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   4
            Top             =   285
            Width           =   1140
         End
         Begin VB.CheckBox Chk_devolucao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(4) Devolução"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   1830
            TabIndex        =   7
            Top             =   285
            Width           =   1410
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Height          =   2055
         Left            =   -74925
         TabIndex        =   36
         Top             =   1320
         Width           =   15195
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14695
            Picture         =   "frm_Natureza_OP.frx":18EB3
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Visualizar arquivo."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14370
            Picture         =   "frm_Natureza_OP.frx":19475
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Limpar caminho."
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton Cmd_valor_pago 
            BackColor       =   &H00C0C0C0&
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
            Left            =   1515
            Picture         =   "frm_Natureza_OP.frx":195B3
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Filtrar por valor pago."
            Top             =   370
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
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Caminho do comprovante."
            Top             =   1590
            Width           =   6705
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14040
            Picture         =   "frm_Natureza_OP.frx":199CE
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Localizar comprovante."
            Top             =   1590
            Width           =   315
         End
         Begin VB.ComboBox txtFormaPagto 
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
            ItemData        =   "frm_Natureza_OP.frx":19AD0
            Left            =   11655
            List            =   "frm_Natureza_OP.frx":19B0A
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            ToolTipText     =   "Forma da baixa."
            Top             =   370
            Width           =   3360
         End
         Begin VB.TextBox Txt_total_juros 
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
            Left            =   7755
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Valor total de juros."
            Top             =   370
            Width           =   1280
         End
         Begin VB.TextBox txt_ValorPago 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Valor baixado."
            Top             =   370
            Width           =   1350
         End
         Begin VB.TextBox Txt_multa 
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
            Left            =   9045
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "Valor da multa."
            Top             =   370
            Width           =   1320
         End
         Begin VB.TextBox Txt_dias_atraso 
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
            Left            =   5130
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Dias em atraso."
            Top             =   370
            Width           =   1245
         End
         Begin VB.TextBox txtobs_pgto 
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
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            ToolTipText     =   "Observações do pagamento."
            Top             =   990
            Width           =   7065
         End
         Begin VB.TextBox txt_Ndocto 
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
            Left            =   13505
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Número do documento baixa."
            Top             =   990
            Width           =   1505
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11940
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "N° da conta corrente."
            Top             =   990
            Width           =   1550
         End
         Begin VB.CommandButton cmdbaixa 
            BackColor       =   &H00C0C0C0&
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
            Left            =   3300
            Picture         =   "frm_Natureza_OP.frx":19C18
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Filtrar por data de pagamento."
            Top             =   370
            Width           =   315
         End
         Begin VB.CommandButton cmdbanco 
            BackColor       =   &H00C0C0C0&
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
            Left            =   11550
            Picture         =   "frm_Natureza_OP.frx":1A033
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Filtrar por instituição bancária."
            Top             =   990
            Width           =   315
         End
         Begin VB.CheckBox chbparcial 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bx. parcial"
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
            Height          =   195
            Left            =   3720
            TabIndex        =   40
            Top             =   420
            Width           =   1335
         End
         Begin VB.TextBox txtjuros 
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
            Left            =   6390
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Valor diário do juros de mora."
            Top             =   370
            Width           =   1350
         End
         Begin VB.TextBox txtdesconto 
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
            Left            =   10380
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Valor do desconto."
            Top             =   370
            Width           =   1260
         End
         Begin VB.ComboBox txtBanco 
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
            Left            =   7320
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   37
            ToolTipText     =   "Instituição bancária."
            Top             =   990
            Width           =   4215
         End
         Begin MSComCtl2.DTPicker txtBaixado 
            Height          =   315
            Left            =   1905
            TabIndex        =   56
            ToolTipText     =   "Data da baixa."
            Top             =   375
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
            Format          =   200998915
            CurrentDate     =   39057
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caminho do comprovante"
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
            Left            =   9757
            TabIndex        =   69
            Top             =   1380
            Width           =   1830
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total de juros"
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
            Left            =   7900
            TabIndex        =   68
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Multa"
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
            Left            =   9510
            TabIndex        =   67
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor baixado"
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
            Left            =   278
            TabIndex        =   66
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias em atraso"
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
            Left            =   5220
            TabIndex        =   65
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Observações do pagamento"
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
            Height          =   240
            Left            =   180
            TabIndex        =   64
            Top             =   780
            Width           =   7065
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
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
            Left            =   9210
            TabIndex        =   63
            Top             =   780
            Width           =   435
         End
         Begin VB.Label LblDocumento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "N° documento baixa"
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
            Left            =   13510
            TabIndex        =   62
            Top             =   780
            Width           =   1505
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conta corrente"
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
            Left            =   12173
            TabIndex        =   61
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. baixa"
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
            Left            =   2227
            TabIndex        =   60
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma da baixa"
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
            Left            =   12780
            TabIndex        =   59
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Juros mora diário"
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
            Left            =   6450
            TabIndex        =   58
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto"
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
            Left            =   10673
            TabIndex        =   57
            Top             =   180
            Width           =   675
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   70
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   12
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
         ButtonCaption8  =   "Validação"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Validar/Cancelar validação (F8)"
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
         ButtonWidth8    =   53
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Atualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Utilizado pelo administrador do sistema."
         ButtonKey9      =   "10"
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
         ButtonLeft9     =   361
         ButtonTop9      =   2
         ButtonWidth9    =   50
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Ajuda"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Ajuda (F1)"
         ButtonKey10     =   "12"
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
         ButtonLeft10    =   413
         ButtonTop10     =   2
         ButtonWidth10   =   36
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonCaption11 =   "Sair"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Sair (Esc)"
         ButtonKey11     =   "13"
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
         ButtonLeft11    =   451
         ButtonTop11     =   2
         ButtonWidth11   =   26
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonKey12     =   "14"
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
         ButtonLeft12    =   479
         ButtonTop12     =   2
         ButtonWidth12   =   24
         ButtonHeight12  =   24
         ButtonUseMaskColor12=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   13770
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frm_Natureza_OP.frx":1A44E
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   88
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
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13770
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frm_Natureza_OP.frx":21CB3
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView lst_NatOp 
         Height          =   3360
         Left            =   75
         TabIndex        =   25
         Top             =   5730
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   5927
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "CFOP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Natureza da operação"
            Object.Width           =   7409
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "ICMS"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "IPI"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "PIS"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "COFINS"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Red. BC ICMS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Soma IPI BC ICMS"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Soma IPI BC ICMS ST"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Soma ret. nos totais"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Object.Tag             =   "T"
            Text            =   "Desc. SUFRAMA"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Object.Tag             =   "T"
            Text            =   "MP aplic."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Object.Tag             =   "T"
            Text            =   "Dest. imp."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Object.Tag             =   "T"
            Text            =   "Validado"
            Object.Width           =   1499
         EndProperty
      End
      Begin MSComctlLib.ListView ListaCliente 
         Height          =   6180
         Left            =   -74925
         TabIndex        =   33
         Top             =   3525
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10901
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Razão social"
            Object.Width           =   23998
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados da CFOP"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   75
         TabIndex        =   75
         Top             =   2370
         Width           =   12165
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Base de cálculo PIS e Cofins"
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
            Left            =   9510
            TabIndex        =   120
            Top             =   720
            Width           =   2445
            Begin VB.CheckBox chkIcmsBasePisCofins 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Abater valor do ICMS"
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
               Left            =   180
               TabIndex        =   121
               Top             =   300
               Width           =   2085
            End
         End
         Begin VB.TextBox txtDtValidacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   375
            Width           =   1635
         End
         Begin VB.TextBox txtRespValidacao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   375
            Width           =   6105
         End
         Begin VB.TextBox mskcfop 
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
            TabIndex        =   20
            ToolTipText     =   "Natureza da operação."
            Top             =   950
            Width           =   855
         End
         Begin VB.TextBox txt_dados_adicionais 
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
            Height          =   1065
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            ToolTipText     =   "Texto padrão para dados adicionais da nota."
            Top             =   1545
            Width           =   3855
         End
         Begin VB.TextBox txt_corpo_nota 
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
            Height          =   1065
            Left            =   4050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            ToolTipText     =   "Texto padrão para corpo da nota."
            Top             =   1545
            Width           =   3855
         End
         Begin VB.TextBox txt_NatOP 
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
            Left            =   1050
            MaxLength       =   50
            TabIndex        =   21
            ToolTipText     =   "Descrição da natureza da operação."
            Top             =   950
            Width           =   8370
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox txtResponsavel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   375
            Width           =   3135
         End
         Begin VB.TextBox Txt_observacoes 
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
            Height          =   1065
            Left            =   7920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            ToolTipText     =   "Observações."
            Top             =   1545
            Width           =   4035
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Index           =   3
            Left            =   4305
            TabIndex        =   106
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   7912
            TabIndex        =   105
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP*"
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
            Left            =   450
            TabIndex        =   82
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Natureza da operação*"
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
            Left            =   4388
            TabIndex        =   81
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Texto dados adicionais (Utilizado na NFe)"
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
            Left            =   630
            TabIndex        =   80
            Top             =   1350
            Width           =   2955
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Texto corpo da nota (Reservado ao Fisco)"
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
            Left            =   4455
            TabIndex        =   79
            Top             =   1350
            Width           =   3045
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Left            =   435
            TabIndex        =   78
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
            Index           =   9
            Left            =   2160
            TabIndex        =   77
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Index           =   1
            Left            =   9465
            TabIndex        =   76
            Top             =   1350
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "frm_Natureza_OP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_CFOP          As Boolean 'OK
Dim Novo_CFOP_Cliente  As Boolean 'OK
Public StrSql_CFOP     As String 'OK
Public FormulaRel_CFOP As String 'OK
Dim TBLISTA_CFOP       As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=jbZRw7tlFpQ&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=23&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If lst_NatOp.ListItems.Count = 0 Then Exit Sub
NomeRel = "Fiscal_CFOP.rpt"
ProcImprimirRel FormulaRel_CFOP, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtid = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao order by IDCountCfop", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("IDCountCfop = " & txtid)
    TBAbrir.MovePrevious
    If TBAbrir.BOF = False Then
        txtid = TBAbrir!IDCountCfop
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & txtid, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        USMsgBox ("Fim dos cadastros de natureza de operação."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior_Cliente()
On Error GoTo tratar_erro

If txtID_CFOP_Cliente = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao_Cliente order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("ID = " & txtID_CFOP_Cliente)
    TBAbrir.MovePrevious
    If TBAbrir.BOF = False Then
        txtID_CFOP_Cliente = TBAbrir!ID
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_NaturezaOperacao_Cliente where ID = " & txtID_CFOP_Cliente, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos_Cliente
        ProcCarregaDados_Cliente
    Else
        USMsgBox ("Fim dos cadastros de cliente."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtid = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao order by IDCountCfop", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("IDCountCfop = " & txtid)
    TBAbrir.MoveNext
    If TBAbrir.EOF = False Then
        txtid = TBAbrir!IDCountCfop
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & txtid, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        USMsgBox ("Fim dos cadastros de natureza de operação."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo_Cliente()
On Error GoTo tratar_erro

If txtID_CFOP_Cliente = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_NaturezaOperacao_Cliente order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("ID = " & txtID_CFOP_Cliente)
    TBAbrir.MoveNext
    If TBAbrir.EOF = False Then
        txtID_CFOP_Cliente = TBAbrir!ID
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_NaturezaOperacao_Cliente where ID = " & txtID_CFOP_Cliente, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos_Cliente
        ProcCarregaDados_Cliente
    Else
        USMsgBox ("Fim dos cadastros de cliente."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnvariavel1_Click()
On Error GoTo tratar_erro

txt_dados_adicionais = txt_dados_adicionais & "@NfVlrTotal"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnvariavel2_Click()
On Error GoTo tratar_erro

txt_dados_adicionais = txt_dados_adicionais & "@NfAliqICMSSN"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnvariavel3_Click()
On Error GoTo tratar_erro

txt_dados_adicionais = txt_dados_adicionais & "@NfVlrICMSSN"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCST_Click()
On Error GoTo tratar_erro

    ProcCST

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_retem_Click()
On Error GoTo tratar_erro

chk_ICMS.Visible = Chk_retem.Value
chk_IPI.Visible = Chk_retem.Value
chk_PIS.Visible = Chk_retem.Value
chk_COFINS.Visible = Chk_retem.Value
chk_Somar.Visible = Chk_retem.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With lst_NatOp
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(9) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(9) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcliente_Click()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CFOP.AbsolutePage <> 2 Then
    If TBLISTA_CFOP.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_CFOP.PageCount - 1)
    Else
        TBLISTA_CFOP.AbsolutePage = TBLISTA_CFOP.AbsolutePage - 2
        ProcExibePagina (TBLISTA_CFOP.AbsolutePage)
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
    TBLISTA_CFOP.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_CFOP.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CFOP.AbsolutePage = 1
ProcExibePagina (TBLISTA_CFOP.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_CFOP.AbsolutePage <> -3 Then
    If TBLISTA_CFOP.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_CFOP.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_CFOP.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_CFOP.AbsolutePage = TBLISTA_CFOP.PageCount
ProcExibePagina (TBLISTA_CFOP.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaCFOP()
On Error GoTo tratar_erro
Var = 6
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM tbl_NaturezaOperacao where IDCountCFOP =" & Var, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

Do While TBAbrir.EOF = False
If TBAbrir!Vendas = True Then
TBAbrir!Tipo_Operacao = 1
End If

If TBAbrir!MaoObra = True Then
TBAbrir!Tipo_Operacao = 2
End If

If TBAbrir!Demonstracao = True Then
TBAbrir!Tipo_Operacao = 3
End If

If TBAbrir!Devolucao = True Then
TBAbrir!Tipo_Operacao = 4
End If

If TBAbrir!Remessa = True Then
TBAbrir!Tipo_Operacao = 5
End If

If TBAbrir!retorno = True Then
TBAbrir!Tipo_Operacao = 6
End If

'Cfop de entrada
If Left(TBAbrir!ID_CFOP, 1) = 1 Or Left(TBAbrir!ID_CFOP, 1) = 2 Or Left(TBAbrir!ID_CFOP, 1) = 3 Then
TBAbrir!Tipo_CFOP = 2
End If

'CFop de saida
If Left(TBAbrir!ID_CFOP, 1) = 4 Or Left(TBAbrir!ID_CFOP, 1) = 5 Or Left(TBAbrir!ID_CFOP, 1) = 6 Then
TBAbrir!Tipo_CFOP = 1
End If

TBAbrir.Update
TBAbrir.MoveNext
Loop
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
    Case "0":
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcGravar
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF5: ProcImprimir
            Case vbKeyF7: ProcCST
            Case vbKeyF8: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros lst_NatOp, "Faturamento/Fiscal/Natureza de operação"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case "1":
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo_Cliente
            Case vbKeyF3: ProcGravar_Cliente
            Case vbKeyF4: ProcExcluir_cliente
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

ProcCarregaToolBar1 Me, 15192, 11, True
ProcCarregaToolBar2 Me, 15192, 10, True

Cmb_opcao_lista = "Validação"
Formulario = "Faturamento/Fiscal/Natureza de operação"
Direitos
SSTab1.Tab = 0
ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me
ProcAtualizaCFOP

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
    
Formulario = "Faturamento/Fiscal/Natureza de operação"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362N" Then frm_Natureza_OP_atualizar.Show 1
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frm_Natureza_OP_atualizar
        If .Chk1.Value = 1 Then
            'Atualiza os dados da CFOP nos dados comerciais dos clientes
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Clientes_DadosComerciais where idCFOP is not null order by idCFOP", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBFI.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBFI.EOF = False
                    If IsNull(TBFI!CFOP) = False And TBFI!CFOP <> "" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select IDCountCfop FROM tbl_NaturezaOperacao WHERE id_CFOP = '" & IIf(IsNull(TBFI!CFOP), 0, TBFI!CFOP) & "' and txt_Descricao = '" & TBFI!descricaoCFOP & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            TBFI!IDCFOP = TBAbrir!IDCountCfop
                            TBFI.Update
                        Else
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select ID_CFOP, Txt_descricao FROM tbl_NaturezaOperacao WHERE IDCountCfop = '" & IIf(IsNull(TBFI!IDCFOP), 0, TBFI!IDCFOP) & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                TBFI!CFOP = TBAbrir!ID_CFOP
                                TBFI!descricaoCFOP = TBAbrir!Txt_descricao
                                TBFI.Update
                            End If
                        End If
                        TBAbrir.Close
                    End If
                    TBFI.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
        
        If .Chk2.Value = 1 Then
            'Excluir CFOP duplicadas
            TextoExcluir = ""
            TextoExcluir1 = ""
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_NaturezaOperacao order by id_CFOP, IDCountCfop", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBFI.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBFI.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select IDCountCfop from tbl_NaturezaOperacao where IDCountCfop > " & TBFI!IDCountCfop & " and id_CFOP = '" & TBFI!ID_CFOP & "' and txt_ICMS = '" & TBFI!Txt_ICMS & "' and txt_IPI = '" & TBFI!txt_IPI & "' and txt_Somar = '" & TBFI!txt_Somar & "' and Retem = " & IIf(TBFI!Retem = True, 1, 0) & " And Proprio = " & IIf(TBFI!Proprio = True, 1, 0) & " And Terceiros = " & IIf(TBFI!Terceiros = True, 1, 0) & " And Suframa = " & IIf(TBFI!Suframa = True, 1, 0) & " And Soma_retorno_totalnf = " & IIf(TBFI!Soma_retorno_totalnf = True, 1, 0) & " And TemPIS = " & IIf(TBFI!TemPIS = True, 1, 0) & " And TemCOFINS = " & IIf(TBFI!TemCOFINS = True, 1, 0) & " And De = " & IIf(TBFI!De = True, 1, 0) & " And FE = " & IIf(TBFI!FE = True, 1, 0) & " And MPA = " & IIf(TBFI!MPA = True, 1, 0) & " And TemReducaoBC = " & IIf(TBFI!TemReducaoBC = True, 1, 0) & " And Somar_IPI_BC_ICMSST = " & IIf(TBFI!Somar_IPI_BC_ICMSST = True, 1, 0), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            Conexao.Execute "UPDATE Clientes_DadosComerciais Set idCFOP = " & TBFI!IDCountCfop & ", CFOP = '" & TBFI!ID_CFOP & "', descricaoCFOP = '" & TBFI!Txt_descricao & "' where idCFOP = " & TBAbrir!IDCountCfop
                            Conexao.Execute "UPDATE Compras_pedido_lista Set ID_CFOP = " & TBFI!IDCountCfop & " where ID_CFOP = " & TBAbrir!IDCountCfop
                            Conexao.Execute "UPDATE tbl_Detalhes_Nota Set ID_CFOP = " & TBFI!IDCountCfop & " where ID_CFOP = " & TBAbrir!IDCountCfop
                            Conexao.Execute "UPDATE vendas_carteira Set ID_CFOP = " & TBFI!IDCountCfop & " where ID_CFOP = " & TBAbrir!IDCountCfop
                            
                            If TextoExcluir = "" Then TextoExcluir = "ID_CFOP = " & TBAbrir!IDCountCfop Else TextoExcluir = TextoExcluir & " or ID_CFOP = " & TBAbrir!IDCountCfop
                            If TextoExcluir1 = "" Then TextoExcluir1 = "IDCountCfop = " & TBAbrir!IDCountCfop Else TextoExcluir1 = TextoExcluir1 & " or IDCountCfop = " & TBAbrir!IDCountCfop
                            TBAbrir.MoveNext
                        Loop
                    End If
                    TBAbrir.Close
                    TBFI.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            Conexao.Execute "DELETE from tbl_NaturezaOperacao_Cliente where " & TextoExcluir
            Conexao.Execute "DELETE from tbl_NaturezaOperacao_CST where " & TextoExcluir
            Conexao.Execute "DELETE from tbl_NaturezaOperacao where " & TextoExcluir1
        End If
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Faturamento/Fiscal/Natureza de operação"
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

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With lst_NatOp
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) natureza(s) de operação?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_NaturezaOperacao WHERE IDCOUNTCFOP = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Faturamento/Fiscal/Natureza de operação"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "CFOP: " & TBFI!ID_CFOP & " - Descrição: " & TBFI!Txt_descricao
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE FROM tbl_NaturezaOperacao WHERE IDCOUNTCFOP = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) natureza(s) de operação antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Natureza(s) de operação excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    Frame9.Enabled = False
    Frame3.Enabled = False
    Frame7.Enabled = False
    Frame4.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    Novo_CFOP = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_cliente()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListaCliente
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) cliente?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_NaturezaOperacao_Cliente WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Faturamento/Fiscal/Natureza de operação"
                Evento = "Excluir cliente"
                ID_documento = .ListItems(InitFor)
                Documento = "CFOP: " & mskcfop & " - Descrição: " & txt_NatOP
                                
                Set TBClientes = CreateObject("adodb.recordset")
                TBClientes.Open "Select * from Clientes where IDCliente = " & IIf(IsNull(TBFI!ID_Cliente), 0, TBFI!ID_Cliente), Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then Documento1 = "ID: " & TBFI!ID_Cliente & " cliente: " & TBClientes!NomeRazao
                TBClientes.Close
                
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE FROM tbl_NaturezaOperacao_Cliente WHERE ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cliente(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Cliente(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos_Cliente
    ProcCarregaLista_Cliente
    Frame8.Enabled = False
    Novo_CFOP_Cliente = False
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
Novo_CFOP = True
Frame9.Enabled = True
Frame3.Enabled = True
Frame7.Enabled = True
Frame4.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
chk_Proprio.SetFocus
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo_Cliente()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("criar novo", txtDtValidacao, "natureza de operação", "cliente", False) = False Then Exit Sub
ProcLimpaCampos_Cliente
Novo_CFOP_Cliente = True
Frame8.Enabled = True
cmdcliente_Click
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_CFOP = True Then
    If USMsgBox("A natureza de operação ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar
        If Novo_CFOP = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_CFOP_Cliente = True Then
    If USMsgBox("O cliente ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcGravar_Cliente
        If Novo_CFOP_Cliente = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_CFOP = False
Novo_CFOP_Cliente = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCliente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaCliente
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtDtValidacao <> "" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaCliente, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCliente_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaCliente
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If txtDtValidacao <> "" Then
                USMsgBox ("Não é permitido excluir este cliente, pois esta natureza de operação está validada."), vbExclamation, "CAPRIND v5.0"
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

Private Sub lst_NatOp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lst_NatOp
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If .ListItems(InitFor).ListSubItems(14) = "SIM" Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    
                    ProcVerificaRegistroUtilizadoSemMsg "Clientes_DadosComerciais", "idCFOP = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "vendas_carteira", "ID_cfop = " & .ListItems(InitFor)
                    If Permitido = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Nota", "ID_cfop = " & .ListItems(InitFor)
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
    ProcOrdenaListView lst_NatOp, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_NatOp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lst_NatOp
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir" Then
                If .ListItems(InitFor).ListSubItems(14) = "SIM" Then
                    USMsgBox ("Não é permitido excluir esta natureza de operação, pois a mesma está validada."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Mensagem = "Não é permitido excluir esta natureza de operação, pois a mesma está sendo utilizada no módulo"
                ProcVerificaRegistroUtilizado "Clientes_DadosComerciais", "idCFOP = " & .ListItems(InitFor), "Vendas/Clientes"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "vendas_carteira", "ID_cfop = " & .ListItems(InitFor), "Vendas/Proposta comercial"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_Detalhes_Nota", "ID_cfop = " & .ListItems(InitFor), "Faturamento/Nota fiscal"
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

Private Sub lst_NatOp_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
 
If lst_NatOp.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_naturezaoperacao where idcountcfop = " & lst_NatOp.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = lst_NatOp.SelectedItem.index
End If
TBLISTA.Close
Frame9.Enabled = True
Frame3.Enabled = True
Frame7.Enabled = True
Frame4.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
    
'ProcVerificaTipoCFOP

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCliente_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
 
If ListaCliente.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_NaturezaOperacao_Cliente where ID = " & ListaCliente.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos_Cliente
    ProcCarregaDados_Cliente
    CodigoLista1 = ListaCliente.SelectedItem.index
End If
TBLISTA.Close
Frame8.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaDados()
On Error GoTo tratar_erro
 
txtid.Text = TBLISTA!IDCountCfop
txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
txtDtValidacao = IIf(IsNull(TBLISTA!DtValidacao), "", TBLISTA!DtValidacao)
txtRespValidacao = IIf(IsNull(TBLISTA!RespValidacao), "", TBLISTA!RespValidacao)
mskcfop.Text = IIf(IsNull(TBLISTA!ID_CFOP), "", TBLISTA!ID_CFOP)
Caption = "Administrativo - Faturamento - Fiscal - Natureza da operação (CFOP : " & TBLISTA!ID_CFOP & " - Natureza : " & TBLISTA!Txt_descricao & ")"
txt_NatOP.Text = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
txt_dados_adicionais.Text = IIf(IsNull(TBLISTA!txt_dados_adicionais), "", TBLISTA!txt_dados_adicionais)
txt_corpo_nota.Text = IIf(IsNull(TBLISTA!txt_corpo_nota), "", TBLISTA!txt_corpo_nota)
txt_observacoes = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
If TBLISTA!Vendas = True Then chkVendas.Value = 1 Else chkVendas.Value = 0
If TBLISTA!MaoObra = True Then chkMaoObra.Value = 1 Else chkMaoObra.Value = 0
If TBLISTA!Demonstracao = True Then Chk_demonstracao.Value = 1 Else Chk_demonstracao.Value = 0
If TBLISTA!Devolucao = True Then Chk_devolucao.Value = 1 Else Chk_devolucao.Value = 0
If TBLISTA!Remessa = True Then Chk_remessa.Value = 1 Else Chk_remessa.Value = 0
If TBLISTA!retorno = True Then Chk_retorno.Value = 1 Else Chk_retorno.Value = 0
If TBLISTA!De = True Then OptDE.Value = True Else OptDE.Value = False
If TBLISTA!FE = True Then optFE.Value = True Else optFE.Value = False
If TBLISTA!Proprio = True Then chk_Proprio.Value = 1 Else chk_Proprio.Value = 0
If TBLISTA!Terceiros = True Then chk_Terceiros.Value = 1 Else chk_Terceiros.Value = 0
If TBLISTA!Txt_ICMS = "SIM" Then chk_ICMS.Value = 1 Else chk_ICMS.Value = 0
If TBLISTA!txt_IPI = "SIM" Then chk_IPI.Value = 1 Else chk_IPI.Value = 0
If TBLISTA!TemPIS = True Then chk_PIS.Value = 1 Else chk_PIS.Value = 0
If TBLISTA!TemCOFINS = True Then chk_COFINS.Value = 1 Else chk_COFINS.Value = 0
If TBLISTA!TemReducaoBC = True Then chkReducao_BC.Value = 1 Else chkReducao_BC.Value = 0
If TBLISTA!CreditaCentroCusto = True Then chkCreditaCentroCusto.Value = 1 Else chkCreditaCentroCusto.Value = 0
If TBLISTA!txt_Somar = "SIM" Then chk_Somar.Value = 1 Else chk_Somar.Value = 0
If TBLISTA!Somar_IPI_BC_ICMSST = True Then Chk_somar_IPI_BC_ICMSST.Value = 1 Else Chk_somar_IPI_BC_ICMSST.Value = 0
If TBLISTA!Soma_retorno_totalnf = True Then Chk_soma_retorno_total.Value = 1 Else Chk_soma_retorno_total.Value = 0
If TBLISTA!Suframa = True Then chkSuframa.Value = 1 Else chkSuframa.Value = 0
If TBLISTA!MPA = True Then Chk_MPA.Value = 1 Else Chk_MPA.Value = 0
If TBLISTA!Retem = True Then Chk_retem.Value = 1 Else Chk_retem.Value = 0
If TBLISTA!Tipo_CFOP = "1" Then optSaida.Value = True Else optEntrada.Value = True

If TBLISTA!AbateICMSBasePisCofins = True Then chkIcmsBasePisCofins.Value = 1 Else chkIcmsBasePisCofins.Value = 0
Novo_CFOP = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados_Cliente()
On Error GoTo tratar_erro
 
txtID_CFOP_Cliente.Text = TBLISTA!ID
txtData_cliente = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel_cliente = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
txtid_cliente = IIf(IsNull(TBLISTA!ID_Cliente), "", TBLISTA!ID_Cliente)
txtDadosAdicionais_cliente.Text = IIf(IsNull(TBLISTA!dados_adicionais), "", TBLISTA!dados_adicionais)
txtCorpoNota_cliente.Text = IIf(IsNull(TBLISTA!Corpo_nota), "", TBLISTA!Corpo_nota)
Novo_CFOP_Cliente = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
lst_NatOp.ListItems.Clear
If StrSql_CFOP = "" Then Exit Sub

'Debug.print StrSql_CFOP
Set TBLISTA_CFOP = CreateObject("adodb.recordset")
TBLISTA_CFOP.Open StrSql_CFOP, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_CFOP.EOF = False Then ProcExibePagina (Pagina)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Cliente()
On Error GoTo tratar_erro

ListaCliente.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select NC.*, C.NomeRazao from tbl_NaturezaOperacao_Cliente NC INNER JOIN Clientes C ON C.IDCliente = NC.ID_Cliente where NC.ID_CFOP = " & IIf(txtid = "", 0, txtid) & " order by C.NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaCliente.ListItems
            .Add = TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!ID_Cliente
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!NomeRazao), "", TBLISTA!NomeRazao)
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

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

lst_NatOp.ListItems.Clear
TBLISTA_CFOP.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_CFOP.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_CFOP.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_CFOP.RecordCount - IIf(Pagina > 1, (TBLISTA_CFOP.PageSize * (Pagina - 1)), 0), TBLISTA_CFOP.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_CFOP.EOF = False And (ContadorReg <= TamanhoPagina)
    With lst_NatOp.ListItems
        .Add , , TBLISTA_CFOP!IDCountCfop
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_CFOP!ID_CFOP), "", TBLISTA_CFOP!ID_CFOP)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_CFOP!Txt_descricao), "", TBLISTA_CFOP!Txt_descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_CFOP!Txt_ICMS), "", TBLISTA_CFOP!Txt_ICMS)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_CFOP!txt_IPI), "", TBLISTA_CFOP!txt_IPI)
        .Item(.Count).SubItems(5) = IIf(TBLISTA_CFOP!TemPIS = True, "SIM", "NÃO")
        .Item(.Count).SubItems(6) = IIf(TBLISTA_CFOP!TemCOFINS = True, "SIM", "NÃO")
        .Item(.Count).SubItems(7) = IIf(TBLISTA_CFOP!TemReducaoBC = True, "SIM", "NÃO")
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_CFOP!txt_Somar), "", TBLISTA_CFOP!txt_Somar)
        .Item(.Count).SubItems(9) = IIf(TBLISTA_CFOP!Somar_IPI_BC_ICMSST = True, "SIM", "NÃO")
        .Item(.Count).SubItems(10) = IIf(TBLISTA_CFOP!Soma_retorno_totalnf = True, "SIM", "NÃO")
        .Item(.Count).SubItems(11) = IIf(TBLISTA_CFOP!Suframa = True, "SIM", "NÃO")
        .Item(.Count).SubItems(12) = IIf(TBLISTA_CFOP!MPA = True, "SIM", "NÃO")
        .Item(.Count).SubItems(13) = IIf(TBLISTA_CFOP!Retem = True, "SIM", "NÃO")
        .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA_CFOP!DtValidacao) = False, "SIM", "NÃO")
    End With
    TBLISTA_CFOP.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_CFOP.RecordCount
If TBLISTA_CFOP.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_CFOP.PageCount
ElseIf TBLISTA_CFOP.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_CFOP.PageCount & " de: " & TBLISTA_CFOP.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_CFOP.AbsolutePage - 1 & " de: " & TBLISTA_CFOP.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frm_Natureza_OP_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtid.Text = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
mskcfop.Text = ""
txt_NatOP.Text = ""
txt_dados_adicionais.Text = ""
txt_corpo_nota.Text = ""
txt_observacoes = ""

chkVendas.Value = 0
chkMaoObra.Value = 0
Chk_demonstracao.Value = 0
Chk_devolucao.Value = 0
Chk_remessa.Value = 0
Chk_retorno.Value = 0

OptDE.Value = False
optFE.Value = False

chk_Proprio.Value = 0
chk_Terceiros.Value = 0

chk_ICMS.Value = 0
chk_IPI.Value = 0
chk_PIS.Value = 0
chk_COFINS.Value = 0
chkReducao_BC.Value = 0
chkCreditaCentroCusto.Value = 0
chk_Somar.Value = 0
Chk_somar_IPI_BC_ICMSST.Value = 0
Chk_soma_retorno_total.Value = 0
chkSuframa.Value = 0
Chk_MPA.Value = 0
Chk_retem.Value = 0

CodigoLista = 0
Caption = "Administrativo - Faturamento - Fiscal - Natureza da operação"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos_Cliente()
On Error GoTo tratar_erro

txtID_CFOP_Cliente.Text = 0
txtid_cliente.Text = 0
txtData_cliente = Format(Date, "dd/mm/yy")
txtResponsavel_cliente = pubUsuario
txtRazao.Text = ""
txtDadosAdicionais_cliente.Text = ""
txtCorpoNota_cliente.Text = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravar()
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
If chk_Proprio.Value = 0 And chk_Terceiros.Value = 0 Then
    NomeCampo = "a aplicação"
    ProcVerificaAcao
    Exit Sub
End If
If OptDE.Value = False And optFE.Value = False Then
    NomeCampo = "o destino"
    ProcVerificaAcao
    Exit Sub
End If
If mskcfop.Text = "" Then
    NomeCampo = "a CFOP"
    ProcVerificaAcao
    If mskcfop.Enabled = True Then mskcfop.SetFocus
    Exit Sub
End If
If Len(mskcfop.Text) < 5 Or Len(mskcfop.Text) > 5 Then
    USMsgBox ("É necessário cadastrar corretamente o número da CFOP - Ex.: 5.101"), vbExclamation, "CAPRIND v5.0"
    mskcfop.SetFocus
    Exit Sub
End If
If txt_NatOP.Text = "" Then
    NomeCampo = "a natureza da operação"
    ProcVerificaAcao
    If txt_NatOP.Enabled = True Then txt_NatOP.SetFocus
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_NaturezaOperacao where IDcountCFOP = " & txtid.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesma", "natureza de operação", False) = False Then Exit Sub
    If txt_NatOP <> TBGravar!Txt_descricao Then Conexao.Execute "Update Clientes_DadosComerciais Set CFOP = '" & mskcfop & "', descricaoCFOP = '" & txt_NatOP & "' where idCFOP = " & TBGravar!IDCountCfop
End If
If txtData <> "" Then TBGravar!Data = txtData Else TBGravar!Data = Date
If txtResponsavel <> "" Then TBGravar!Responsavel = txtResponsavel Else TBGravar!Responsavel = pubUsuario
TBGravar!ID_CFOP = mskcfop.Text
TBGravar!Txt_descricao = txt_NatOP.Text
TBGravar!txt_dados_adicionais = txt_dados_adicionais.Text
TBGravar!txt_corpo_nota = txt_corpo_nota.Text
TBGravar!Obs = txt_observacoes

If chkVendas.Value = 1 Then
TBGravar!Vendas = True
TBGravar!Tipo_Operacao = "1"
Else
TBGravar!Vendas = False
End If

If chkMaoObra.Value = 1 Then
TBGravar!MaoObra = True
TBGravar!Tipo_Operacao = "2"
Else
TBGravar!MaoObra = False
End If

If Chk_demonstracao.Value = 1 Then
TBGravar!Demonstracao = True
TBGravar!Tipo_Operacao = "3"
Else
TBGravar!Demonstracao = False
End If

If Chk_devolucao.Value = 1 Then
TBGravar!Devolucao = True
TBGravar!Tipo_Operacao = "4"
Else
TBGravar!Devolucao = False
End If

If Chk_remessa.Value = 1 Then
TBGravar!Remessa = True
TBGravar!Tipo_Operacao = "5"
Else
TBGravar!Remessa = False
End If

If Chk_retorno.Value = 1 Then
TBGravar!retorno = True
TBGravar!Tipo_Operacao = "6"
Else
TBGravar!retorno = False
End If

If OptDE.Value = True Then TBGravar!De = True Else TBGravar!De = False
If optFE.Value = True Then TBGravar!FE = True Else TBGravar!FE = False

If optEntrada.Value = True Then TBGravar!Tipo_CFOP = "2"
If optSaida.Value = True Then TBGravar!Tipo_CFOP = "1"

If chk_Proprio.Value = 1 Then TBGravar!Proprio = 1 Else TBGravar!Proprio = 0
If chk_Terceiros.Value = 1 Then TBGravar!Terceiros = 1 Else TBGravar!Terceiros = 0
If chk_ICMS.Value = 1 Then TBGravar!Txt_ICMS = "SIM" Else TBGravar!Txt_ICMS = "NÃO"
If chk_IPI.Value = 1 Then TBGravar!txt_IPI = "SIM" Else TBGravar!txt_IPI = "NÃO"
If chk_PIS.Value = 1 Then TBGravar!TemPIS = True Else TBGravar!TemPIS = False
If chk_COFINS.Value = 1 Then TBGravar!TemCOFINS = True Else TBGravar!TemCOFINS = False
If chkReducao_BC.Value = 1 Then TBGravar!TemReducaoBC = True Else TBGravar!TemReducaoBC = False
If chk_Somar.Value = 1 Then TBGravar!txt_Somar = "SIM" Else TBGravar!txt_Somar = "NÃO"
If Chk_somar_IPI_BC_ICMSST.Value = 1 Then TBGravar!Somar_IPI_BC_ICMSST = True Else TBGravar!Somar_IPI_BC_ICMSST = False
If Chk_soma_retorno_total.Value = 1 Then TBGravar!Soma_retorno_totalnf = True Else TBGravar!Soma_retorno_totalnf = False
If chkSuframa.Value = 1 Then TBGravar!Suframa = True Else TBGravar!Suframa = False
If Chk_MPA.Value = 1 Then TBGravar!MPA = True Else TBGravar!MPA = False
If Chk_retem.Value = 1 Then TBGravar!Retem = True Else TBGravar!Retem = False
If chkCreditaCentroCusto = 1 Then TBGravar!CreditaCentroCusto = True Else TBGravar!CreditaCentroCusto = False
' Retira ICMS da base de calculo do PI e Cofins
If chkIcmsBasePisCofins.Value = 1 Then TBGravar!AbateICMSBasePisCofins = True Else TBGravar!AbateICMSBasePisCofins = False

TBGravar.Update
txtid = TBGravar!IDCountCfop
TBGravar.Close
If Novo_CFOP = True Then
    USMsgBox ("Nova natureza de operação cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    StrSql_CFOP = "Select * from tbl_NaturezaOperacao where IDcountCFOP = " & txtid.Text
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And lst_NatOp.ListItems.Count <> 0 Then
        lst_NatOp.SelectedItem = lst_NatOp.ListItems(CodigoLista)
        lst_NatOp.SetFocus
    End If
End If
1:
    '==================================
    Modulo = "Faturamento/Fiscal/Natureza de operação"
    ID_documento = txtid
    Documento = "CFOP: " & mskcfop & " - Descrição: " & txt_NatOP
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_CFOP = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravar_Cliente()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "natureza de operação", "cliente", False) = False Then Exit Sub
Acao = "salvar"
If txtid_cliente.Text = "0" Or txtid_cliente.Text = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    cmdcliente_Click
    Exit Sub
End If

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from tbl_NaturezaOperacao_Cliente where ID_CFOP = " & txtid.Text & " and ID_cliente = " & txtid_cliente & " and ID <> " & txtID_CFOP_Cliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    USMsgBox ("Já existe cadastro deste cliente para esta CFOP."), vbExclamation, "CAPRIND v5.0"
    TBClientes.Close
    Exit Sub
End If
TBClientes.Close

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_NaturezaOperacao_Cliente where ID = " & txtID_CFOP_Cliente.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Data = IIf(txtData_cliente = "", Date, txtData_cliente)
TBGravar!Responsavel = IIf(txtResponsavel_cliente = "", pubUsuario, txtResponsavel_cliente)
TBGravar!ID_CFOP = txtid
TBGravar!ID_Cliente = txtid_cliente
TBGravar!dados_adicionais = txtDadosAdicionais_cliente.Text
TBGravar!Corpo_nota = txtCorpoNota_cliente.Text
TBGravar.Update
txtID_CFOP_Cliente = TBGravar!ID
TBGravar.Close

ProcCarregaLista_Cliente
If Novo_CFOP_Cliente = True Then
    USMsgBox ("Cliente cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo cliente"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar cliente"
    If ListaCliente.ListItems.Count <> 0 And CodigoLista1 <> 0 Then
        ListaCliente.SelectedItem = ListaCliente.ListItems(CodigoLista1)
        ListaCliente.SetFocus
    End If
End If
'==================================
Modulo = "Faturamento/Fiscal/Natureza de operação"
ID_documento = txtid
Documento = "CFOP: " & mskcfop & " - Descrição: " & txt_NatOP
Documento1 = "ID: " & txtid_cliente & " cliente: " & txtRazao
ProcGravaEvento
'==================================
Novo_CFOP_Cliente = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCST()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_CFOP = True Then
    USMsgBox ("Salve a CFOP antes de cadastrar as CST."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "cadastrar as CST"
If txtid = 0 Then
    NomeCampo = "a CFOP"
    ProcVerificaAcao
    Exit Sub
End If
frm_Natureza_OP_CST.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtid = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        If lst_NatOp.Visible = True Then lst_NatOp.SetFocus
    Case 1:
        If Novo_CFOP = True Then
            USMsgBox ("Salve a natureza de operação antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
            SSTab1.Tab = 0
            Exit Sub
        End If
        ListaCliente.SetFocus
        ProcCarregaLista_Cliente
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub ProcVerificaTipoCFOP()
On Error GoTo tratar_erro


Texto = txt_NatOP.Text
USMsgBox Left$(Texto, 5)
'If InStr(1, Texto, "Venda", vbTextCompare) > 0 Then

If Left$(Texto, 5) = "VENDA" Then
USMsgBox "Venda"
chkVendas.Value = 1
End If

If Left$(Texto, 9) = "DEVOLUCAO" Or Left$(Texto, 9) = "DEVOLUÇÃO" Then
USMsgBox Left$(Texto, 9)
Chk_devolucao.Value = 1
End If

If InStr(1, Texto, "Remessa", vbTextCompare) > 0 Then
USMsgBox "Remessa"
Chk_remessa.Value = 1
End If

If InStr(1, Texto, "Retorno", vbTextCompare) > 0 Then
USMsgBox "Retorno"
Chk_retorno.Value = 1
End If

If InStr(1, Texto, "INDUSTRIALIZAÇÃO", vbTextCompare) > 0 Then
USMsgBox "INDUSTRIALIZAÇÃO"
chkMaoObra.Value = 1
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtID_Change()
On Error GoTo tratar_erro

If txtid.Text = "" Then Exit Sub

Lista.ListItems.Clear

Set TBCST = CreateObject("adodb.recordset")
TBCST.Open "Select * from tbl_NaturezaOperacao_CST where ID_CFOP = " & txtid.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBCST.EOF = False Then
    Do While TBCST.EOF = False
        With Lista.ListItems
            .Add , , TBCST!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBCST!CST_ICMS), "", TBCST!CST_ICMS)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBCST!CST_IPI), "", TBCST!CST_IPI)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBCST!CST_PIS), "", TBCST!CST_PIS)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBCST!CST_Cofins), "", TBCST!CST_Cofins)
        End With
        TBCST.MoveNext
    Loop
End If
TBCST.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtID_cliente_Change()
On Error GoTo tratar_erro

txtRazao = ""
If txtid_cliente <> "" Then
    VerifNumero = txtid_cliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtid_cliente = ""
        txtid_cliente.SetFocus
        Exit Sub
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Clientes where IDCliente = " & txtid_cliente, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        txtRazao = IIf(IsNull(TBFI!NomeRazao), "", (TBFI!NomeRazao))
    End If
    TBFI.Close
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcGravar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcValidarRegistros lst_NatOp, "Faturamento/Fiscal/Natureza de operação"
    Case 9: ProcAtualizar
    Case 10: ProcAjuda
    Case 11: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo_Cliente
    Case 2: ProcGravar_Cliente
    Case 3: ProcExcluir_cliente
    Case 4: ProcImprimir
    Case 5: ProcAnterior_Cliente
    Case 6: ProcProximo_Cliente
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
