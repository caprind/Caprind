VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Vendedores 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Vendedores"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   1425
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      ScreenHeightDT  =   768
      ScreenWidthDT   =   1024
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5685
      Left            =   45
      TabIndex        =   19
      Top             =   3390
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   10028
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Nº vend."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Vendedor"
         Object.Width           =   20602
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Comissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Validado"
         Object.Width           =   1499
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   92
      Top             =   9720
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar2 
      Height          =   975
      Left            =   60
      TabIndex        =   93
      Top             =   330
      Visible         =   0   'False
      Width           =   15255
      _ExtentX        =   26908
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
      ButtonLeft4     =   118
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   122
      ButtonTop5      =   2
      ButtonWidth5    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
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
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   11970
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Vendedores.frx":0000
         Count           =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmVendas_Vendedores.frx":33D8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "USToolBar1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Clientes"
      TabPicture(1)   =   "frmVendas_Vendedores.frx":33F4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIDVendedor_Cliente"
      Tab(1).Control(1)=   "Lista2"
      Tab(1).Control(2)=   "Frame11"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Produtos/serviços"
      TabPicture(2)   =   "frmVendas_Vendedores.frx":3410
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "txtIDProduto"
      Tab(2).Control(2)=   "txtIDVendedor_Prod"
      Tab(2).Control(3)=   "Lista3"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Clientes/produtos/serviços"
      TabPicture(3)   =   "frmVendas_Vendedores.frx":342C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Comissões por meta"
      TabPicture(4)   =   "frmVendas_Vendedores.frx":3448
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lista6"
      Tab(4).Control(1)=   "FrameMeta"
      Tab(4).Control(2)=   "USToolBar3"
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissões"
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
         Height          =   2055
         Left            =   12945
         TabIndex        =   64
         Top             =   1290
         Width           =   2355
         Begin VB.CheckBox Chk_bloquear_venda 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bloquear venda outros clientes"
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
            Height          =   375
            Left            =   90
            TabIndex        =   117
            Top             =   1530
            Width           =   2175
         End
         Begin VB.OptionButton optMeta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meta mensal (R$)"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   105
            Top             =   585
            Width           =   2175
         End
         Begin VB.OptionButton optVendedor 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vendedor"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   15
            Top             =   345
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optCliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cliente"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   16
            Top             =   825
            Width           =   2175
         End
         Begin VB.OptionButton optProduto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Produto/serviço"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   17
            Top             =   1050
            Width           =   2175
         End
         Begin VB.OptionButton optCliente_prod 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cliente/Produto/serviço"
            BeginProperty Font 
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
            Left            =   60
            TabIndex        =   18
            Top             =   1290
            Width           =   2175
         End
      End
      Begin MSComctlLib.ListView Lista6 
         Height          =   7245
         Left            =   -74970
         TabIndex        =   107
         Top             =   2400
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   12779
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Valor de"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Valor até"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Comissão de (%)"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Frame FrameMeta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão por meta"
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
         Height          =   1095
         Left            =   -74970
         TabIndex        =   108
         Top             =   1290
         Width           =   15255
         Begin VB.TextBox txtIDMeta 
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
            Left            =   9660
            MaxLength       =   50
            TabIndex        =   116
            ToolTipText     =   "Número do vendedor."
            Top             =   510
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtMcomissao 
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
            Left            =   6990
            MaxLength       =   50
            TabIndex        =   113
            ToolTipText     =   "Número do vendedor."
            Top             =   510
            Width           =   1125
         End
         Begin VB.TextBox txtAte 
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
            Left            =   3450
            MaxLength       =   50
            TabIndex        =   111
            ToolTipText     =   "Número do vendedor."
            Top             =   510
            Width           =   1125
         End
         Begin VB.TextBox txtDe 
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
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   109
            ToolTipText     =   "Número do vendedor."
            Top             =   510
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
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
            Left            =   8220
            TabIndex        =   115
            Top             =   600
            Width           =   165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pagar comissão de"
            BeginProperty Font 
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
            Left            =   5580
            TabIndex        =   114
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até"
            BeginProperty Font 
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
            Left            =   3150
            TabIndex        =   112
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Do Valor de"
            BeginProperty Font 
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
            Left            =   330
            TabIndex        =   110
            Top             =   600
            Width           =   825
         End
      End
      Begin MSComctlLib.ListView Lista3 
         Height          =   6975
         Left            =   -74925
         TabIndex        =   33
         Top             =   2730
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   12303
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
         NumItems        =   5
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
            Object.Width           =   12938
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Comissão"
            Object.Width           =   1764
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   91
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
         ButtonLeft5     =   156
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
         ButtonLeft6     =   205
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Validação"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Validar/Cancelar validação (F10)"
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
         ButtonLeft7     =   253
         ButtonTop7      =   2
         ButtonWidth7    =   53
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
         ButtonLeft8     =   308
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
         ButtonLeft9     =   312
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
         ButtonLeft10    =   350
         ButtonTop10     =   2
         ButtonWidth10   =   26
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
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
         ButtonLeft11    =   378
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   6870
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_Vendedores.frx":3464
            Count           =   1
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   75
         TabIndex        =   94
         Top             =   9120
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
            TabIndex        =   74
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
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
            ItemData        =   "frmVendas_Vendedores.frx":9407
            Left            =   6960
            List            =   "frmVendas_Vendedores.frx":9411
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   210
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
            TabIndex        =   76
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   80
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Vendedores.frx":9429
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
            TabIndex        =   79
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Vendedores.frx":CBD0
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
            TabIndex        =   77
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
            TabIndex        =   78
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Vendedores.frx":106DF
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
            TabIndex        =   81
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Vendedores.frx":147D2
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
            Left            =   3360
            TabIndex        =   104
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Index           =   30
            Left            =   2040
            TabIndex        =   101
            Top             =   240
            Width           =   645
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
            Index           =   0
            Left            =   5610
            TabIndex        =   100
            Top             =   240
            Width           =   1260
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
            Left            =   6960
            TabIndex        =   99
            Top             =   270
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
            TabIndex        =   96
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
            TabIndex        =   95
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.TextBox txtIDVendedor_Prod 
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
         Left            =   -74310
         MaxLength       =   50
         MouseIcon       =   "frmVendas_Vendedores.frx":1805F
         MousePointer    =   99  'Custom
         TabIndex        =   87
         Text            =   "0"
         ToolTipText     =   "Número do vendedor."
         Top             =   4260
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtIDProduto 
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
         Left            =   -74370
         MaxLength       =   50
         MouseIcon       =   "frmVendas_Vendedores.frx":18369
         MousePointer    =   99  'Custom
         TabIndex        =   86
         Text            =   "0"
         ToolTipText     =   "Número do vendedor."
         Top             =   3540
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtIDVendedor_Cliente 
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
         Left            =   -74250
         MaxLength       =   50
         MouseIcon       =   "frmVendas_Vendedores.frx":18673
         MousePointer    =   99  'Custom
         TabIndex        =   85
         Text            =   "0"
         ToolTipText     =   "Número do vendedor."
         Top             =   3180
         Visible         =   0   'False
         Width           =   825
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8715
         Left            =   -74925
         TabIndex        =   66
         Top             =   1290
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15372
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
         TabCaption(0)   =   "Clientes"
         TabPicture(0)   =   "frmVendas_Vendedores.frx":1897D
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtIDVendedor_ClienteProd"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Produtos/serviços"
         TabPicture(1)   =   "frmVendas_Vendedores.frx":18999
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista5"
         Tab(1).Control(1)=   "Frame14"
         Tab(1).Control(2)=   "txtIDVendedor_ClienteProd2"
         Tab(1).Control(3)=   "txtIDProduto_ClienteProd"
         Tab(1).ControlCount=   4
         Begin VB.TextBox txtIDProduto_ClienteProd 
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
            Left            =   -74580
            MaxLength       =   50
            MouseIcon       =   "frmVendas_Vendedores.frx":189B5
            MousePointer    =   99  'Custom
            TabIndex        =   90
            Text            =   "0"
            Top             =   2820
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox txtIDVendedor_ClienteProd2 
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
            Left            =   -73500
            MaxLength       =   50
            MouseIcon       =   "frmVendas_Vendedores.frx":18CBF
            MousePointer    =   99  'Custom
            TabIndex        =   89
            Text            =   "0"
            ToolTipText     =   "Número do vendedor."
            Top             =   2790
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox txtIDVendedor_ClienteProd 
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
            Left            =   1020
            MaxLength       =   50
            MouseIcon       =   "frmVendas_Vendedores.frx":18FC9
            MousePointer    =   99  'Custom
            TabIndex        =   88
            Text            =   "0"
            ToolTipText     =   "Número do vendedor."
            Top             =   2160
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Frame Frame13 
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
            Height          =   795
            Left            =   30
            TabIndex        =   82
            Top             =   330
            Width           =   15165
            Begin VB.TextBox txtCliente_prod 
               BackColor       =   &H00FFFFFF&
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
               Left            =   915
               MaxLength       =   255
               TabIndex        =   35
               ToolTipText     =   "Razão social."
               Top             =   390
               Width           =   8655
            End
            Begin VB.TextBox txtIDCliente_prod 
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
               TabIndex        =   34
               ToolTipText     =   "Código do cliente."
               Top             =   390
               Width           =   720
            End
            Begin VB.TextBox txtCidadeCliente_prod 
               BackColor       =   &H00FFFFFF&
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
               Left            =   9990
               MaxLength       =   150
               TabIndex        =   37
               ToolTipText     =   "Cidade."
               Top             =   390
               Width           =   4965
            End
            Begin VB.CommandButton cmdCliente_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   9585
               Picture         =   "frmVendas_Vendedores.frx":192D3
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Localizar produtos/itens."
               Top             =   390
               Width           =   315
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Razão social"
               BeginProperty Font 
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
               Left            =   4800
               TabIndex        =   84
               Top             =   180
               Width           =   885
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
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
               Index           =   7
               Left            =   12255
               TabIndex        =   83
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.Frame Frame14 
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
            Height          =   1425
            Left            =   -74970
            TabIndex        =   69
            Top             =   330
            Width           =   15165
            Begin VB.CommandButton cmdDesenhoCliente_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   3840
               Picture         =   "frmVendas_Vendedores.frx":193D5
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Localizar produtos/serviços."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtDescricaoCliente_prod 
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
               MaxLength       =   255
               TabIndex        =   44
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   990
               Width           =   13320
            End
            Begin VB.TextBox txtDesenhoCliente_prod 
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
               TabIndex        =   39
               ToolTipText     =   "Código interno."
               Top             =   390
               Width           =   2775
            End
            Begin VB.TextBox txtFamiliaCliente_prod 
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
               Left            =   4230
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   43
               TabStop         =   0   'False
               ToolTipText     =   "Família."
               Top             =   390
               Width           =   10740
            End
            Begin VB.CommandButton cmdFiltrarCliente_prod 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   3510
               Picture         =   "frmVendas_Vendedores.frx":194D7
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Filtrar por código interno."
               Top             =   390
               Width           =   315
            End
            Begin VB.TextBox txtRevCliente_prod 
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
               Left            =   2970
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   40
               TabStop         =   0   'False
               Text            =   "0"
               ToolTipText     =   "Revisão do produto/item."
               Top             =   390
               Width           =   525
            End
            Begin VB.TextBox txtComissaoCliente_prod 
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
               Left            =   13515
               MaxLength       =   50
               TabIndex        =   45
               ToolTipText     =   "Comissão do vendedor (%)."
               Top             =   990
               Width           =   1455
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
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
               Left            =   6495
               TabIndex        =   73
               Top             =   780
               Width           =   690
            End
            Begin VB.Label Label15 
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
               Left            =   9375
               TabIndex        =   72
               Top             =   180
               Width           =   480
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
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
               Left            =   1035
               TabIndex        =   71
               Top             =   180
               Width           =   1230
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Comissão"
               BeginProperty Font 
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
               Left            =   13905
               TabIndex        =   70
               Top             =   780
               Width           =   675
            End
         End
         Begin MSComctlLib.ListView Lista5 
            Height          =   6615
            Left            =   -74970
            TabIndex        =   46
            Top             =   1770
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   11668
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
            NumItems        =   5
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
               Object.Width           =   12832
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Família"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Comissão"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2055
         Left            =   75
         TabIndex        =   48
         Top             =   1290
         Width           =   12855
         Begin VB.TextBox txtMeta 
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
            Left            =   11310
            MaxLength       =   50
            TabIndex        =   122
            ToolTipText     =   "Valor da meta mensal"
            Top             =   1560
            Width           =   1395
         End
         Begin VB.TextBox txtComissao 
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
            Left            =   9900
            MaxLength       =   50
            TabIndex        =   120
            ToolTipText     =   "Comissão do vendedor (%)."
            Top             =   1560
            Width           =   1395
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
            ItemData        =   "frmVendas_Vendedores.frx":198F2
            Left            =   180
            List            =   "frmVendas_Vendedores.frx":198F4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   118
            ToolTipText     =   "Empresa."
            Top             =   390
            Width           =   3285
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
            Height          =   315
            Left            =   180
            MaxLength       =   255
            TabIndex        =   10
            ToolTipText     =   "E-mail."
            Top             =   1560
            Width           =   3645
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   2970
            Top             =   150
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Cmd_visualizar_arquivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9540
            Picture         =   "frmVendas_Vendedores.frx":198F6
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Visualizar arquivo."
            Top             =   1560
            Width           =   315
         End
         Begin VB.CommandButton Cmd_limpar_caminho 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9210
            Picture         =   "frmVendas_Vendedores.frx":19EB8
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpar caminho."
            Top             =   1560
            Width           =   315
         End
         Begin VB.CommandButton cmdImportar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   8880
            Picture         =   "frmVendas_Vendedores.frx":19FF6
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Localizar assinatura."
            Top             =   1560
            Width           =   315
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
            Left            =   8190
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1605
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
            Left            =   9810
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   2910
         End
         Begin VB.TextBox txtnumero 
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
            Left            =   3480
            MaxLength       =   50
            TabIndex        =   0
            ToolTipText     =   "Número do vendedor."
            Top             =   390
            Width           =   675
         End
         Begin VB.TextBox txtvendedor 
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
            Left            =   4170
            MaxLength       =   255
            TabIndex        =   1
            ToolTipText     =   "Nome do vendedor."
            Top             =   390
            Width           =   4005
         End
         Begin VB.TextBox txtregiao 
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
            Left            =   4980
            MaxLength       =   255
            TabIndex        =   5
            ToolTipText     =   "Região."
            Top             =   960
            Width           =   2085
         End
         Begin VB.TextBox txttelefone1 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   6
            ToolTipText     =   "Telefone."
            Top             =   960
            Width           =   1405
         End
         Begin VB.TextBox txtfax 
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
            Left            =   11325
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   "Fax."
            Top             =   960
            Width           =   1395
         End
         Begin VB.TextBox txttelefone2 
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
            Left            =   8505
            MaxLength       =   50
            TabIndex        =   7
            ToolTipText     =   "Telefone."
            Top             =   960
            Width           =   1395
         End
         Begin VB.TextBox txttelefone3 
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
            Left            =   9915
            MaxLength       =   50
            TabIndex        =   8
            ToolTipText     =   "Telefone."
            Top             =   960
            Width           =   1395
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
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Caminho da assinatura."
            Top             =   1560
            Width           =   5025
         End
         Begin VB.TextBox txtendereco 
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   255
            TabIndex        =   4
            ToolTipText     =   "Endereço."
            Top             =   960
            Width           =   4785
         End
         Begin VB.Label lblMeta 
            BackStyle       =   0  'Transparent
            Caption         =   "Meta mensal R$"
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
            Left            =   11490
            TabIndex        =   123
            Top             =   1350
            Width           =   1155
         End
         Begin VB.Label lblComissao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão (%)"
            BeginProperty Font 
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
            Left            =   10095
            TabIndex        =   121
            Top             =   1350
            Width           =   1005
         End
         Begin VB.Label Label1 
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
            Index           =   14
            Left            =   1455
            TabIndex        =   119
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho da assinatura"
            BeginProperty Font 
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
            Left            =   5535
            TabIndex        =   103
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº vend."
            BeginProperty Font 
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
            Left            =   3540
            TabIndex        =   102
            Top             =   180
            Width           =   645
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável validação"
            BeginProperty Font 
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
            Left            =   10448
            TabIndex        =   98
            Top             =   180
            Width           =   1635
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Index           =   2
            Left            =   8310
            TabIndex        =   97
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor"
            BeginProperty Font 
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
            Left            =   5820
            TabIndex        =   56
            Top             =   180
            Width           =   690
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Região"
            BeginProperty Font 
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
            Left            =   5775
            TabIndex        =   55
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
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
            Left            =   2235
            TabIndex        =   54
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel."
            BeginProperty Font 
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
            Left            =   7642
            TabIndex        =   53
            Top             =   750
            Width           =   270
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. 04"
            BeginProperty Font 
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
            Left            =   11775
            TabIndex        =   52
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
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
            Left            =   1867
            TabIndex        =   51
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. 02"
            BeginProperty Font 
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
            Left            =   8955
            TabIndex        =   50
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. 03"
            BeginProperty Font 
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
            Left            =   10365
            TabIndex        =   49
            Top             =   750
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView Lista2 
         Height          =   7545
         Left            =   -74925
         TabIndex        =   25
         Top             =   2130
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   13309
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Razão social"
            Object.Width           =   20346
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cidade"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Comissão"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame Frame11 
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
         Height          =   825
         Left            =   -74925
         TabIndex        =   58
         Top             =   1290
         Width           =   15225
         Begin VB.TextBox txtComissaoCliente 
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
            Left            =   13500
            MaxLength       =   50
            TabIndex        =   24
            ToolTipText     =   "Comissão vendedor (%)."
            Top             =   390
            Width           =   1535
         End
         Begin VB.CommandButton cmdCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9855
            Picture         =   "frmVendas_Vendedores.frx":1A0F8
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Localizar cliente."
            Top             =   390
            Width           =   315
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
            Left            =   10260
            MaxLength       =   150
            TabIndex        =   23
            ToolTipText     =   "Cidade."
            Top             =   390
            Width           =   3225
         End
         Begin VB.TextBox txtidCliente 
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
            TabIndex        =   20
            ToolTipText     =   "Código do cliente."
            Top             =   390
            Width           =   720
         End
         Begin VB.TextBox txtnomerazao 
            BackColor       =   &H00FFFFFF&
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
            Left            =   915
            MaxLength       =   255
            TabIndex        =   21
            ToolTipText     =   "Razão social."
            Top             =   390
            Width           =   8925
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão"
            BeginProperty Font 
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
            Left            =   13930
            TabIndex        =   67
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   11625
            TabIndex        =   65
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Razão social"
            BeginProperty Font 
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
            Left            =   4935
            TabIndex        =   59
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1395
         Left            =   -74925
         TabIndex        =   60
         Top             =   1320
         Width           =   15225
         Begin VB.TextBox txtComissaoProduto 
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
            Left            =   13485
            MaxLength       =   50
            TabIndex        =   32
            ToolTipText     =   "Comissão do vendedor (%)."
            Top             =   990
            Width           =   1515
         End
         Begin VB.TextBox txtRev_cod 
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
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Revisão do produto/item."
            Top             =   390
            Width           =   525
         End
         Begin VB.CommandButton cmdFiltrarProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3510
            Picture         =   "frmVendas_Vendedores.frx":1A1FA
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Filtrar por código interno."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtFamilia 
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
            Left            =   4230
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   390
            Width           =   10800
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
            MaxLength       =   50
            TabIndex        =   26
            ToolTipText     =   "Código interno."
            Top             =   390
            Width           =   2775
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
            MaxLength       =   255
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   990
            Width           =   13290
         End
         Begin VB.CommandButton cmdProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   3840
            Picture         =   "frmVendas_Vendedores.frx":1A615
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Localizar produtos/serviços."
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão"
            BeginProperty Font 
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
            Left            =   13905
            TabIndex        =   68
            Top             =   780
            Width           =   675
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   1035
            TabIndex        =   63
            Top             =   180
            Width           =   1230
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
            Left            =   9390
            TabIndex        =   62
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
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
            Left            =   6480
            TabIndex        =   61
            Top             =   780
            Width           =   690
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74970
         TabIndex        =   106
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
         ButtonLeft4     =   118
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft5     =   122
         ButtonTop5      =   2
         ButtonWidth5    =   36
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   160
         ButtonTop6      =   2
         ButtonWidth6    =   26
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonKey7      =   "7"
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
         ButtonLeft7     =   188
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         ButtonUseMaskColor7=   0   'False
         Begin DrawSuite2022.USImageList USImageList3 
            Left            =   11970
            Top             =   240
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmVendas_Vendedores.frx":1A717
            Count           =   1
         End
      End
      Begin VB.TextBox txtID 
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
         Left            =   1560
         MaxLength       =   50
         MouseIcon       =   "frmVendas_Vendedores.frx":1DAEF
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Text            =   "0"
         ToolTipText     =   "Número do vendedor."
         Top             =   1680
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin MSComctlLib.ListView Lista4 
      Height          =   7245
      Left            =   105
      TabIndex        =   38
      Top             =   2400
      Visible         =   0   'False
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   12779
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Razão social"
         Object.Width           =   22005
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cidade"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmVendas_Vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_vendedores     As Boolean 'OK
Dim Novo_vendedores1    As Boolean 'OK
Dim Novo_vendedores2    As Boolean 'OK
Dim Novo_vendedores3    As Boolean 'OK
Dim Novo_vendedores4    As Boolean 'OK
Dim AlterarComissao     As Boolean 'OK
Public Sql_Vendedores   As String 'OK
Dim TBLISTA_Vendedores  As ADODB.Recordset 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=PayU1hae_7E&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=58&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtNumero.Text = ""
txtVendedor.Text = ""
txtDtValidacao = ""
txtRespValidacao = ""
txtComissao.Text = ""
txtendereco.Text = ""
txtregiao.Text = ""
txtFax.Text = ""
txttelefone1.Text = ""
txttelefone2.Text = ""
txttelefone3.Text = ""
txtEmail.Text = ""
txt_Caminho = ""
Chk_bloquear_venda.Value = 0
CodigoLista = 0
Caption = "Administrativo - Vendas - Vendedores"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBVendas!n_vendedor = txtNumero.Text
TBVendas!vendedor = txtVendedor.Text
TBVendas!Comissao = IIf(txtComissao.Text = "", Null, txtComissao.Text)
TBVendas!regiao = txtregiao.Text
TBVendas!Endereco = IIf(txtendereco.Text = "", Null, txtendereco.Text)
TBVendas!telefone1 = txttelefone1.Text
TBVendas!telefone2 = txttelefone2.Text
TBVendas!telefone3 = txttelefone3.Text
TBVendas!Fax = txtFax.Text
TBVendas!Email = IIf(txtEmail.Text = "", Null, LCase(txtEmail.Text))
TBVendas!Assinatura = txt_Caminho
If Chk_bloquear_venda.Value = 1 Then TBVendas!Bloquear_venda_cliente = True Else TBVendas!Bloquear_venda_cliente = False
Texto = ""
If optVendedor = True Then Texto = "V"
If optCliente = True Then Texto = "C"
If optProduto = True Then Texto = "P"
If optCliente_prod = True Then Texto = "CP"
If optMeta = True Then Texto = "MT"
If TBVendas!tipocomissao <> Texto Then
    Conexao.Execute "delete from vendas_vendedores_clientes where IDVendedor = " & txtId
    Conexao.Execute "delete from vendas_vendedores_produto where IDVendedor = " & txtId
End If
TBVendas!tipocomissao = Texto
TBVendas!Meta = IIf(txtMeta.Text <> "", txtMeta, 0)
TBVendas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista(Pagina As Integer)
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If Sql_Vendedores = "" Then Exit Sub
Set TBLISTA_Vendedores = CreateObject("adodb.recordset")
TBLISTA_Vendedores.Open Sql_Vendedores, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA_Vendedores.EOF = False Then ProcExibePagina (Pagina)

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
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(7) = 5
    Else
        .ButtonState(4) = 5
        .ButtonState(7) = 0
    End If
    .Refresh
End With

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

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txt_Caminho <> "" Then ProcAbrirArquivo txt_Caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdcliente_Click()
On Error GoTo tratar_erro

Aplic = 1
ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdCliente_prod_Click()
On Error GoTo tratar_erro

Aplic = 2
ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False
frmVendas_LocalizarCliente.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenhoCliente_prod_Click()
On Error GoTo tratar_erro

Aplic = 2
frmVendas_vendedores_produto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluirCliente()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) cliente(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from vendas_vendedores_clientes WHERE id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Vendedores"
            Evento = "Excluir cliente"
            ID_documento = .ListItems(InitFor)
            Documento = "Vendedor: " & txtVendedor
            Documento1 = "Cliente: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cliente(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Cliente(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCliente
    ProcCarregalistaCliente
    Frame11.Enabled = False
    Novo_vendedores1 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluirMeta()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista6
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir essa(s) meta(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from Vendas_Vendedores_Comissoes WHERE id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Vendedores"
            Evento = "Excluir meta"
            ID_documento = .ListItems(InitFor)
            Documento = "Vendedor: " & txtVendedor
            Documento1 = "Meta: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) meta(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Meta(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    txtDe.Text = ""
    txtAte.Text = ""
    txtMcomissao.Text = ""
    txtIDMeta.Text = 0
    
    ProcCarregalistaMetas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procExcluir_ProdCli()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "mesmo", "o vendedor", "alterar", True, True) = False Then Exit Sub
Select Case SSTab2.Tab
    Case 0: procExcluirCliente_ProdCli
    Case 1: procExcluirProduto_ProdCli
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirProduto()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista3
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from vendas_vendedores_produto WHERE id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Vendedores"
            Evento = "Excluir produto/serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Vendedor: " & txtVendedor
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposProduto
    ProcCarregalistaProduto
    Frame5.Enabled = False
    Novo_vendedores2 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrarCliente_prod_Click()
On Error GoTo tratar_erro

If txtDesenhoCliente_prod <> "" Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "select * from projproduto where desenho = '" & txtDesenhoCliente_prod & "' and Vendas = 'True' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        txtIDProduto_ClienteProd = TBItem!Codproduto
        procPuxaProduto2
    End If
    TBItem.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrarProduto_Click()
On Error GoTo tratar_erro

If txtdesenho <> "" Then
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "select * from projproduto where desenho = '" & txtdesenho & "' and Vendas = 'True' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        txtidproduto = TBItem!Codproduto
        procPuxaProduto
    End If
    TBItem.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoCliente()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "cliente", "criar novo", True, True) = False Then Exit Sub
ProcLimpaCamposCliente
Novo_vendedores1 = True
Frame11.Enabled = True
cmdcliente_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovo_ProdCli()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0:
        If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "cliente", "criar novo", True, True) = False Then Exit Sub
        ProcLimpaCamposCliente_prod
        Novo_vendedores3 = True
        Frame13.Enabled = True
        cmdCliente_prod_Click
    Case 1:
        If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "produto/serviço", "criar novo", True, True) = False Then Exit Sub
        ProcLimpaCamposCliente_prod2
        Novo_vendedores4 = True
        Frame14.Enabled = True
        txtDesenhoCliente_prod.SetFocus
End Select

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
If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "produto/serviço", "criar novo", True, True) = False Then Exit Sub
ProcLimpaCamposProduto
Novo_vendedores2 = True
Frame5.Enabled = True
txtdesenho.SetFocus

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

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendedores.AbsolutePage <> 2 Then
    If TBLISTA_Vendedores.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Vendedores.PageCount - 1)
    Else
        TBLISTA_Vendedores.AbsolutePage = TBLISTA_Vendedores.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Vendedores.AbsolutePage)
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
    TBLISTA_Vendedores.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Vendedores.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendedores.AbsolutePage = 1
ProcExibePagina (TBLISTA_Vendedores.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendedores.AbsolutePage <> -3 Then
    If TBLISTA_Vendedores.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Vendedores.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Vendedores.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendedores.AbsolutePage = TBLISTA_Vendedores.PageCount
ProcExibePagina (TBLISTA_Vendedores.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProduto_Click()
On Error GoTo tratar_erro

Aplic = 1
frmVendas_vendedores_produto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvarCliente()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame11.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtIDcliente.Text = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    If Frame11.Enabled = True Then cmdcliente_Click
    Exit Sub
End If
If txtComissaoCliente.Text = "" Then
    NomeCampo = "a comissão"
    ProcVerificaAcao
    If Frame11.Enabled = True Then txtComissaoCliente.SetFocus
    Exit Sub
Else
    VerifNumero = txtComissaoCliente.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissaoCliente.Text = ""
        If Frame11.Enabled = True Then txtComissaoCliente.SetFocus
        Exit Sub
    End If
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_vendedores_clientes where id = " & txtIDVendedor_Cliente.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "o cliente", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDvendedor = txtId
TBGravar!IDCliente = txtIDcliente
TBGravar!Comissao = txtComissaoCliente
TBGravar!tipocomissao = "C"
TBGravar.Update
txtIDVendedor_Cliente = TBGravar!ID
TBGravar.Close
ProcCarregalistaCliente
If Novo_vendedores1 = True Then
    USMsgBox ("Novo cliente cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo cliente"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar cliente"
    If CodigoLista1 <> 0 And Lista2.ListItems.Count <> 0 Then
        Lista2.SelectedItem = Lista2.ListItems(CodigoLista1)
        Lista2.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Vendedores"
ID_documento = txtIDVendedor_Cliente
Documento = "Vendedor: " & txtVendedor
Documento1 = "Cliente: " & txtnomerazao
ProcGravaEvento
'==================================
Novo_vendedores1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvar_ProdCli()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Select Case SSTab2.Tab
    Case 0: ProcSalvarCliente_ProdCli
    Case 1: procSalvarProduto_ProdCli
End Select

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
If Frame5.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtdesenho.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    If Frame5.Enabled = True Then frmVendas_vendedores_produto.Show 1
    Exit Sub
End If
If txtComissaoProduto.Text = "" Then
    NomeCampo = "a comissão"
    ProcVerificaAcao
    If Frame5.Enabled = True Then txtComissaoProduto.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_vendedores_produto where id = " & txtIDVendedor_Prod.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "o produto/serviço", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDvendedor = txtId
TBGravar!IDProduto = txtidproduto
TBGravar!Comissao = txtComissaoProduto
TBGravar.Update
txtIDVendedor_Prod = TBGravar!ID
TBGravar.Close
ProcCarregalistaProduto
If Novo_vendedores2 = True Then
    USMsgBox ("Novo produto/serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo cliente"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar cliente"
    If CodigoLista2 <> 0 And Lista3.ListItems.Count <> 0 Then
        Lista3.SelectedItem = Lista3.ListItems(CodigoLista2)
        Lista3.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Vendedores"
ID_documento = txtIDVendedor_Prod
Documento = "Vendedor: " & txtVendedor
Documento1 = "Cód. interno: " & txtdesenho
ProcGravaEvento
'==================================
Novo_vendedores2 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Vendedores"
Direitos
ProcLimpaVariaveisPrincipais

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
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_vendedores", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBAbrir.MoveLast
    txtNumero.Text = TBAbrir!n_vendedor + 1
Else
    txtNumero.Text = 1
End If
Novo_vendedores = True
optVendedor.Value = True
Frame1.Enabled = True
Frame6.Enabled = True
With txtNumero
    .Locked = False
    .TabStop = True
    .SetFocus
End With
ProcLimparTudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparTudo()
On Error GoTo tratar_erro

Frame11.Enabled = False
Frame5.Enabled = False
Frame13.Enabled = False
Frame14.Enabled = False
ProcLimpaCamposCliente
ProcLimpaCamposProduto
ProcLimpaCamposCliente_prod
ProcLimpaCamposCliente_prod2
Novo_vendedores1 = False
Novo_vendedores2 = False
Novo_vendedores3 = False
Novo_vendedores4 = False

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
                If USMsgBox("Deseja realmente excluir este(s) vendedor(es)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from vendas_vendedores WHERE id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from vendas_vendedores_clientes WHERE IDVendedor = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from vendas_vendedores_produto WHERE IDVendedor = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Vendedores"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Vendedor: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) vendedor(es) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Vendedor(es) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcLimparTudo
    ProcAtualizalista (1)
    Frame1.Enabled = False
    Frame6.Enabled = False
    Novo_vendedores = False
End If

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
Acao = "salvar"
If txtNumero.Text = "" Then
    NomeCampo = "o número do vendedor"
    ProcVerificaAcao
    If Frame1.Enabled = True And txtNumero.Locked = False Then txtNumero.SetFocus
    Exit Sub
End If
If Novo_vendedores = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ID from Vendas_Vendedores where N_Vendedor = " & txtNumero, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Numero do vendedor já cadastrado."), vbExclamation, "CAPRIND v5.0"
        If Frame1.Enabled = True And txtNumero.Locked = False Then txtNumero.SetFocus
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If
If txtVendedor.Text = "" Then
    NomeCampo = "o nome do vendedor"
    ProcVerificaAcao
    If Frame1.Enabled = True Then txtVendedor.SetFocus
    Exit Sub
End If
If optVendedor.Value = True And txtComissao.Text = "" Then
    NomeCampo = "a comissão do vendedor"
    ProcVerificaAcao
    If Frame1.Enabled = True Then txtComissao.SetFocus
    Exit Sub
End If
If txtregiao.Text = "" Then
    NomeCampo = "a região do vendedor"
    ProcVerificaAcao
    If Frame1.Enabled = True Then txtregiao.SetFocus
    Exit Sub
End If
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_vendedores where id = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = True Then
    TBVendas.AddNew
Else
    If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "mesmo", "o vendedor", "alterar", True, True) = False Then Exit Sub
End If
ProcEnviaDados
TBVendas.Update
txtId = TBVendas!ID

If Novo_vendedores = True Then
    USMsgBox ("Novo vendedor cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    Sql_Vendedores = "Select * from vendas_vendedores where id = " & txtId.Text
    ProcAtualizalista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Vendedores"
ID_documento = txtId
Documento = "Vendedor: " & txtVendedor
Documento1 = ""
ProcGravaEvento
'==================================
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_vendedores where id = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    ProcPuxaDados
End If
TBVendas.Close

Novo_vendedores = False
AlterarComissao = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 11, True
ProcCarregaToolBar2 Me, 15195, 7, True
ProcCarregaToolBar3 Me, 15195, 7, True
ProcCarregaComboEmpresa Cmb_empresa, False

USToolBar2.Visible = False
Formulario = "Vendas/Vendedores"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Cmb_opcao_lista = "Validação"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmVendas_vendedores_localizar.Show 1

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
                If Cmb_opcao_lista = "Excluir" Then If FunVerificaRegistroValidadoSemMsg("vendas_vendedores", "id = " & .ListItems(InitFor), True) = False Then GoTo Proximo
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
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("vendas_vendedores", "id = " & .ListItems(InitFor), "mesmo", "este vendedor", "excluir", True, True) = False Then
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
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_vendedores where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    Frame1.Enabled = True
    Frame6.Enabled = True
    txtNumero.Locked = True
    txtNumero.TabStop = False
    CodigoLista = Lista.SelectedItem.index
End If
TBVendas.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId.Text = 0 Then Exit Sub
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_vendedores order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.BOF = False Then
    TBVendas.Find ("ID = " & txtId)
    TBVendas.MovePrevious
    If TBVendas.BOF = False Then
        txtId = TBVendas!ID
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from vendas_vendedores where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
        ProcCarregalistaCliente
        ProcCarregalistaProduto
        ProcCarregalistaClienteProd
    Else
        USMsgBox ("Fim dos cadastros de vendedores."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_vendedores = False
Novo_vendedores1 = False
Novo_vendedores2 = False
Novo_vendedores3 = False
Novo_vendedores4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId.Text = 0 Then Exit Sub
Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from vendas_vendedores order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.BOF = False Then
    TBVendas.Find ("ID = " & txtId)
    TBVendas.MoveNext
    If TBVendas.EOF = False Then
        txtId = TBVendas!ID
        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from vendas_vendedores where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcPuxaDados
        ProcCarregalistaCliente
        ProcCarregalistaProduto
        ProcCarregalistaClienteProd
    Else
        USMsgBox ("Fim dos cadastros de vendedores."), vbInformation, "CAPRIND v5.0"
    End If
End If
Novo_vendedores = False
Novo_vendedores1 = False
Novo_vendedores2 = False
Novo_vendedores3 = False
Novo_vendedores4 = False

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
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
            Case vbKeyF10: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros Lista, "Vendas/Vendedores"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoCliente
            Case vbKeyF3: procSalvarCliente
            Case vbKeyF4: procExcluirCliente
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoProduto
            Case vbKeyF3: ProcSalvarProduto
            Case vbKeyF4: ProcExcluirProduto
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: procNovo_ProdCli
            Case vbKeyF3: procSalvar_ProdCli
            Case vbKeyF4: procExcluir_ProdCli
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoMeta
            Case vbKeyF3: ProcSalvarMeta
            Case vbKeyF4: procExcluirMeta
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
        
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista2
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("vendas_vendedores", "id = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista2, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista2
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "este cliente", "excluir", True, True) = False Then
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

Private Sub Lista2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista2.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_vendedores_clientes where ID = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposCliente
    txtIDVendedor_Cliente = TBAbrir!ID
    txtIDcliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
    txtComissaoCliente = IIf(IsNull(TBAbrir!Comissao), "", TBAbrir!Comissao)
    procPuxaCliente
    Frame11.Enabled = True
    CodigoLista1 = Lista2.SelectedItem.index
End If
TBAbrir.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista3
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("vendas_vendedores", "id = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista3, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista3
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "este produto/serviço", "excluir", True, True) = False Then
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

Private Sub Lista3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista3.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_vendedores_produto where ID = " & Lista3.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposProduto
    txtIDVendedor_Prod = TBAbrir!ID
    txtidproduto = IIf(IsNull(TBAbrir!IDProduto), "", TBAbrir!IDProduto)
    txtComissaoProduto = IIf(IsNull(TBAbrir!Comissao), "", TBAbrir!Comissao)
    procPuxaProduto
    Frame5.Enabled = True
    CodigoLista2 = Lista3.SelectedItem.index
End If
TBAbrir.Close
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista4
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("vendas_vendedores", "id = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista4, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista4
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "este cliente", "excluir", True, True) = False Then
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

Private Sub Lista4_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista4.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_vendedores_clientes where ID = " & Lista4.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposCliente_prod
    txtIDVendedor_ClienteProd = TBAbrir!ID
    txtIDCliente_prod = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
    procPuxaCliente2
    Frame13.Enabled = True
    CodigoLista3 = Lista4.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista5
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("vendas_vendedores", "id = " & txtId, True) = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista5, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista5_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista5
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "este produto/serviço", "excluir", True, True) = False Then
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

Private Sub Lista5_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista5.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_vendedores_produto where ID = " & Lista5.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposCliente_prod2
    txtIDVendedor_ClienteProd2 = TBAbrir!ID
    txtIDProduto_ClienteProd = IIf(IsNull(TBAbrir!IDProduto), "", TBAbrir!IDProduto)
    txtComissaoCliente_prod = IIf(IsNull(TBAbrir!Comissao), "", TBAbrir!Comissao)
    procPuxaProduto2
    Frame14.Enabled = True
    CodigoLista4 = Lista5.SelectedItem.index
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista6_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista6.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_Vendedores_Comissoes where ID = " & Lista6.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtIDMeta = TBAbrir!ID
    txtDe = IIf(IsNull(TBAbrir!De), "", Format(TBAbrir!De, "###,##0.00"))
    txtAte = IIf(IsNull(TBAbrir!Ate), "", Format(TBAbrir!Ate, "###,##0.00"))
    txtMcomissao = IIf(IsNull(TBAbrir!Comissao), "", Format(TBAbrir!Comissao, "###,##0.00"))
    CodigoLista6 = Lista6.SelectedItem.index
    FrameMeta.Enabled = True
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optCliente_Click()
On Error GoTo tratar_erro

If optCliente.Value = True Then
    With txtComissao
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    Chk_bloquear_venda.Enabled = True
    AlterarComissao = True
    'Frame3.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optCliente_prod_Click()
On Error GoTo tratar_erro

If optCliente_prod.Value = True Then
    With txtComissao
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    Chk_bloquear_venda.Enabled = True
    AlterarComissao = True
    'Frame3.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optMeta_Click()
On Error GoTo tratar_erro

'If optMeta.Value = True Then
'    With txtComissao
'        lblComissao.Visible = False
'        .Visible = False
'        .Text = ""
'        .Locked = False
'        .TabStop = True
'        Frame3.Visible = True
'        txtMeta.Visible = True
'        lblMeta.Visible = True
'    End With
'    With Chk_bloquear_venda
'        .Value = 0
'        .Enabled = True
'    End With
'    AlterarComissao = True
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optProduto_Click()
On Error GoTo tratar_erro

If optProduto.Value = True Then
    With txtComissao
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With Chk_bloquear_venda
        .Value = 0
        .Enabled = False
    End With
    AlterarComissao = True
    'Frame3.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optVendedor_Click()
On Error GoTo tratar_erro

If optVendedor.Value = True Then
    With txtComissao
        lblComissao.Visible = True
        .Visible = True
        .Text = ""
        .Locked = False
        .TabStop = True
    End With
    With Chk_bloquear_venda
        .Value = 0
        .Enabled = False
    End With
    AlterarComissao = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtId.Text = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        USToolBar2.Visible = False
        Lista4.Visible = False
        Lista.Visible = True
        If Lista.Visible = True Then Lista.SetFocus
    Case 1:
        Lista4.Visible = False
        Lista.Visible = False
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista2.SetFocus
        ProcLimpaCamposCliente
        ProcCarregalistaCliente
    Case 2:
        Lista4.Visible = False
        Lista.Visible = False
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista3.SetFocus
        ProcLimpaCamposProduto
        ProcCarregalistaProduto
    Case 3:
        Lista4.Visible = True
        Lista.Visible = False
        USToolBar2.Visible = True
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        SSTab2.Tab = 0
        Lista4.SetFocus
        ProcLimpaCamposCliente_prod
        ProcCarregalistaClienteProd
    Case 4:
        Lista4.Visible = False
        Lista.Visible = False
        USToolBar2.Visible = False
        ProcCarregalistaMetas
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtIDVendedor_ClienteProd.Text = "0" Or Novo_vendedores3 = True Then
    SSTab2.Tab = 0
    Exit Sub
End If

Select Case SSTab2.Tab
    Case 0:
        Lista4.Visible = True
        If Lista4.Visible = True Then Lista4.SetFocus
    Case 1:
        Lista4.Visible = False
        Lista5.SetFocus
        ProcLimpaCamposCliente_prod2
        ProcCarregalistaClienteProd2
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtAte_Change()
On Error GoTo tratar_erro

If txtAte.Text <> "" Then
    VerifNumero = txtAte.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtAte.Text = ""
        txtAte.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtAte_LostFocus()
On Error GoTo tratar_erro

txtAte.Text = Format(txtAte.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtcomissao_Change()
On Error GoTo tratar_erro

If txtComissao.Text <> "" Then
    VerifNumero = txtComissao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissao.Text = ""
        txtComissao.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissao_LostFocus()
On Error GoTo tratar_erro

txtComissao.Text = Format(txtComissao.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoCliente_Change()
On Error GoTo tratar_erro

If txtComissaoCliente.Text <> "" Then
    VerifNumero = txtComissaoCliente.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissaoCliente.Text = ""
        txtComissaoCliente.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoCliente_LostFocus()
On Error GoTo tratar_erro

txtComissaoCliente.Text = Format(txtComissaoCliente.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoCliente_prod_Change()
On Error GoTo tratar_erro

If txtComissaoCliente_prod.Text <> "" Then
    VerifNumero = txtComissaoCliente_prod.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissaoCliente_prod.Text = ""
        txtComissaoCliente_prod.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoCliente_prod_LostFocus()
On Error GoTo tratar_erro

txtComissaoCliente_prod.Text = Format(txtComissaoCliente_prod.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoProduto_Change()
On Error GoTo tratar_erro

If txtComissaoProduto.Text <> "" Then
    VerifNumero = txtComissaoProduto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtComissaoProduto.Text = ""
        txtComissaoProduto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissaoProduto_LostFocus()
On Error GoTo tratar_erro

txtComissaoProduto.Text = Format(txtComissaoProduto.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtDe_Change()
On Error GoTo tratar_erro

If txtDe.Text <> "" Then
    VerifNumero = txtDe.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDe.Text = ""
        txtDe.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDe_LostFocus()
On Error GoTo tratar_erro

txtDe.Text = Format(txtDe.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

txtidproduto = ""
txtRev_cod = ""
txtfamilia = ""
txtdescricao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenhoCliente_prod_Change()
On Error GoTo tratar_erro

txtIDProduto_ClienteProd = ""
txtRevCliente_prod = ""
txtFamiliaCliente_prod = ""
txtDescricaoCliente_prod = ""

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

Private Sub txtIDCliente_LostFocus()
On Error GoTo tratar_erro

If txtIDcliente <> "" Then
    VerifNumero = txtIDcliente
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDcliente = ""
        txtIDcliente.SetFocus
        Exit Sub
    End If
    procPuxaCliente
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDCliente_prod_LostFocus()
On Error GoTo tratar_erro

If txtIDCliente_prod <> "" Then
    VerifNumero = txtIDCliente_prod
    ProcVerificaNumero
    If VerifNumero = False Then
        txtIDCliente_prod = ""
        txtIDCliente_prod.SetFocus
        Exit Sub
    End If
    procPuxaCliente2
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMcomissao_Change()
On Error GoTo tratar_erro

If txtMcomissao <> "" Then
    VerifNumero = txtMcomissao
    ProcVerificaNumero
    If VerifNumero = False Then
        txtMcomissao = ""
        txtMcomissao.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMcomissao_LostFocus()
On Error GoTo tratar_erro

txtMcomissao.Text = Format(txtMcomissao.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMeta_Change()
On Error GoTo tratar_erro

If txtMeta <> "" Then
    VerifNumero = txtMeta
    ProcVerificaNumero
    If VerifNumero = False Then
        txtMeta = ""
        txtMeta.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMeta_LostFocus()
On Error GoTo tratar_erro

txtMeta.Text = Format(txtMeta.Text, "###,##0.00")

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

Private Sub txtNumero_LostFocus()
On Error GoTo tratar_erro

If txtNumero.Text <> "" Then
    VerifNumero = txtNumero.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero.Text = ""
        txtNumero.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Caption = "Administrativo - Vendas - Vendedores (Vendedor : " & IIf(IsNull(TBVendas!vendedor), "", TBVendas!vendedor) & ")"
txtId.Text = TBVendas!ID
txtNumero.Text = IIf(IsNull(TBVendas!n_vendedor), "", TBVendas!n_vendedor)
txtVendedor.Text = IIf(IsNull(TBVendas!vendedor), "", TBVendas!vendedor)
txtDtValidacao = IIf(IsNull(TBVendas!DtValidacao), "", TBVendas!DtValidacao)
txtRespValidacao = IIf(IsNull(TBVendas!RespValidacao), "", TBVendas!RespValidacao)
txtendereco.Text = IIf(IsNull(TBVendas!Endereco), "", TBVendas!Endereco)
txtFax.Text = IIf(IsNull(TBVendas!Fax), "", TBVendas!Fax)
txttelefone1.Text = IIf(IsNull(TBVendas!telefone1), "", TBVendas!telefone1)
txttelefone2.Text = IIf(IsNull(TBVendas!telefone2), "", TBVendas!telefone2)
txttelefone3.Text = IIf(IsNull(TBVendas!telefone3), "", TBVendas!telefone3)
txtregiao.Text = IIf(IsNull(TBVendas!regiao), "", TBVendas!regiao)
txtEmail.Text = IIf(IsNull(TBVendas!Email), "", TBVendas!Email)
txt_Caminho = IIf(IsNull(TBVendas!Assinatura), "", TBVendas!Assinatura)

If IsNull(TBVendas!ID_empresa) = False And TBVendas!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBVendas!ID_empresa


If TBVendas!Bloquear_venda_cliente = True Then Chk_bloquear_venda.Value = 1 Else Chk_bloquear_venda.Value = 0
If TBVendas!tipocomissao = "V" Then optVendedor.Value = True
If TBVendas!tipocomissao = "C" Then optCliente.Value = True
If TBVendas!tipocomissao = "P" Then optProduto.Value = True
If TBVendas!tipocomissao = "CP" Then optCliente_prod.Value = True
If TBVendas!tipocomissao = "MT" Then optMeta.Value = True

procComissao
txtComissao.Text = IIf(IsNull(TBVendas!Comissao), "", Format(TBVendas!Comissao, "###,##0.00"))
txtMeta.Text = IIf(IsNull(TBVendas!Meta), "", Format(TBVendas!Meta, "###,##0.00"))

Novo_vendedores = False
AlterarComissao = False
ProcLimparTudo
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procComissao()
On Error GoTo tratar_erro

If optVendedor.Value = True Then
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = False
    SSTab1.TabsPerRow = 1
End If
If optCliente.Value = True Then
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = False
    SSTab1.TabsPerRow = 2
End If
If optProduto.Value = True Then
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = True
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = False
    SSTab1.TabsPerRow = 2
End If
If optCliente_prod.Value = True Then
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = True
    SSTab1.TabVisible(4) = False
    SSTab1.TabsPerRow = 2
End If
If optMeta.Value = True Then
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(4) = True
    SSTab1.TabsPerRow = 2
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposCliente()
On Error GoTo tratar_erro

txtIDVendedor_Cliente = 0
txtIDcliente.Text = ""
txtnomerazao.Text = ""
txtCidade.Text = ""
txtComissaoCliente.Text = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxaCliente()
On Error GoTo tratar_erro

txtnomerazao = ""
txtCidade = ""
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "select * from clientes where idcliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    txtnomerazao = TBClientes!NomeRazao
    txtCidade = TBClientes!Cidade
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxaCliente2()
On Error GoTo tratar_erro

txtCliente_prod = ""
txtCidadeCliente_prod = ""
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "select * from clientes where idcliente = " & txtIDCliente_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    txtCliente_prod = TBClientes!NomeRazao
    txtCidadeCliente_prod = TBClientes!Cidade
End If
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxaProduto()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select Codproduto, Desenho, RevDesenho, classe, Descricao from Projproduto where codproduto = " & txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtdesenho = TBProduto!Desenho
    txtidproduto = TBProduto!Codproduto
    txtRev_cod = TBProduto!RevDesenho
    txtfamilia = TBProduto!Classe
    txtdescricao = TBProduto!Descricao
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procPuxaProduto2()
On Error GoTo tratar_erro

Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "select Codproduto, Desenho, RevDesenho, classe, Descricao from Projproduto where codproduto = " & txtIDProduto_ClienteProd, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    txtDesenhoCliente_prod = TBProduto!Desenho
    txtIDProduto_ClienteProd = TBProduto!Codproduto
    txtRevCliente_prod = TBProduto!RevDesenho
    txtFamiliaCliente_prod = TBProduto!Classe
    txtDescricaoCliente_prod = TBProduto!Descricao
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposProduto()
On Error GoTo tratar_erro

txtIDVendedor_Prod = 0
txtidproduto.Text = 0
txtRev_cod.Text = 0
txtdesenho.Text = ""
txtfamilia.Text = ""
txtdescricao.Text = ""
txtComissaoProduto.Text = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalistaProduto()
On Error GoTo tratar_erro

Lista3.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_vendedores_produto where IDVendedor = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista3.ListItems
            .Add , , TBLISTA!ID
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "select * from Projproduto where codproduto = " & TBLISTA!IDProduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                .Item(.Count).SubItems(1) = TBProduto!Desenho
                .Item(.Count).SubItems(2) = TBProduto!Descricao
                .Item(.Count).SubItems(3) = TBProduto!Classe
            End If
            TBProduto.Close
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Comissao), "", Format(TBLISTA!Comissao, "###,##0.00") & "%")
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

Sub ProcCarregalistaCliente()
On Error GoTo tratar_erro

Lista2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_vendedores_clientes where IDVendedor = " & txtId & " and TipoComissao = 'C'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista2.ListItems
            .Add , , TBLISTA!ID
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "select * from clientes where idcliente = " & TBLISTA!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                .Item(.Count).SubItems(1) = TBClientes!NomeRazao
                .Item(.Count).SubItems(2) = TBClientes!Cidade
            End If
            TBClientes.Close
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Comissao), "", TBLISTA!Comissao & "%")
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

Sub ProcCarregalistaMetas()
On Error GoTo tratar_erro

Lista6.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from Vendas_Vendedores_Comissoes where IDVendedor = " & txtId & " ORDER by De", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista6.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!De), "", Format(TBLISTA!De, "###,##0.00"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ate), "", Format(TBLISTA!Ate, "###,##0.00"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Comissao), "", Format(TBLISTA!Comissao, "###,##0.00") & "%")
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSair()
On Error GoTo tratar_erro

If Novo_vendedores = True Then
    If USMsgBox("O vendedor ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_vendedores = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_vendedores1 = True Then
    If USMsgBox("O cliente ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvarCliente
        If Novo_vendedores1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_vendedores2 = True Then
    If USMsgBox("O produto/serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarProduto
        If Novo_vendedores2 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_vendedores3 = True Then
    If USMsgBox("O cliente ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarCliente_ProdCli
        If Novo_vendedores3 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_vendedores4 = True Then
    If USMsgBox("O produto/serviço ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        procSalvarProduto_ProdCli
        If Novo_vendedores4 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_vendedores = False
Novo_vendedores1 = False
Novo_vendedores2 = False
Novo_vendedores3 = False
Novo_vendedores4 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposCliente_prod()
On Error GoTo tratar_erro

txtIDVendedor_ClienteProd = 0
txtIDCliente_prod.Text = ""
txtCliente_prod.Text = ""
txtCidadeCliente_prod.Text = ""
CodigoLista3 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposCliente_prod2()
On Error GoTo tratar_erro

txtIDVendedor_ClienteProd2 = 0
txtIDProduto_ClienteProd.Text = 0
txtDesenhoCliente_prod.Text = ""
txtRevCliente_prod = 0
txtFamiliaCliente_prod.Text = ""
txtDescricaoCliente_prod.Text = ""
txtComissaoCliente_prod.Text = ""
CodigoLista4 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarCliente_ProdCli()
On Error GoTo tratar_erro

If Frame13.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtIDCliente_prod.Text = "" Then
    NomeCampo = "o cliente"
    ProcVerificaAcao
    If Frame13.Enabled = True Then cmdCliente_prod_Click
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_vendedores_clientes where id = " & txtIDVendedor_ClienteProd.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "o cliente", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDvendedor = txtId
TBGravar!IDCliente = txtIDCliente_prod
TBGravar!tipocomissao = "CP"
TBGravar.Update
txtIDVendedor_ClienteProd = TBGravar!ID
TBGravar.Close
ProcCarregalistaClienteProd
If Novo_vendedores3 = True Then
    USMsgBox ("Novo cliente cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo cliente"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar cliente"
    If CodigoLista3 <> 0 And Lista4.ListItems.Count <> 0 Then
        Lista4.SelectedItem = Lista4.ListItems(CodigoLista3)
        Lista4.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Vendedores"
ID_documento = txtIDVendedor_ClienteProd
Documento = "Vendedor: " & txtVendedor
Documento1 = "Cliente: " & txtCliente_prod
ProcGravaEvento
'==================================
Novo_vendedores3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procSalvarProduto_ProdCli()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame14.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtDesenhoCliente_prod.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    If Frame14.Enabled = True Then cmdDesenhoCliente_prod_Click
    Exit Sub
End If
If txtComissaoCliente_prod.Text = "" Then
    NomeCampo = "a comissão"
    ProcVerificaAcao
    If Frame14.Enabled = True Then txtComissaoCliente_prod.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_vendedores_produto where id = " & txtIDVendedor_ClienteProd2.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "o produto/serviço", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDvendedor = txtId
TBGravar!IDProduto = txtIDProduto_ClienteProd
TBGravar!IDCliente = txtIDVendedor_ClienteProd
TBGravar!Comissao = txtComissaoCliente_prod
TBGravar.Update
txtIDVendedor_ClienteProd2 = TBGravar!ID
TBGravar.Close
ProcCarregalistaClienteProd2
If Novo_vendedores4 = True Then
    USMsgBox ("Novo produto/serviço cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo cliente"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar cliente"
    If CodigoLista4 <> 0 And Lista5.ListItems.Count <> 0 Then
        Lista5.SelectedItem = Lista5.ListItems(CodigoLista4)
        Lista5.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Vendedores"
ID_documento = txtIDVendedor_ClienteProd2
Documento = "Vendedor: " & txtVendedor
Documento1 = "Cód. interno: " & txtDesenhoCliente_prod
ProcGravaEvento
'==================================
Novo_vendedores4 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalistaClienteProd()
On Error GoTo tratar_erro

Lista4.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_vendedores_clientes where IDVendedor = " & txtId & " and TipoComissao = 'CP'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista4.ListItems
            .Add , , TBLISTA!ID
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "select * from clientes where idcliente = " & TBLISTA!IDCliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                .Item(.Count).SubItems(1) = TBClientes!NomeRazao
                .Item(.Count).SubItems(2) = TBClientes!Cidade
            End If
            TBClientes.Close
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

Sub ProcCarregalistaClienteProd2()
On Error GoTo tratar_erro

Lista5.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from vendas_vendedores_produto where IDcliente = " & txtIDVendedor_ClienteProd, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista5.ListItems
            .Add , , TBLISTA!ID
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "select * from Projproduto where codproduto = " & TBLISTA!IDProduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                .Item(.Count).SubItems(1) = TBProduto!Desenho
                .Item(.Count).SubItems(2) = TBProduto!Descricao
                .Item(.Count).SubItems(3) = TBProduto!Classe
            End If
            TBProduto.Close
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Comissao), "", Format(TBLISTA!Comissao, "###,##0.00") & "%")
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

Sub procExcluirCliente_ProdCli()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista4
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) cliente(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from vendas_vendedores_clientes WHERE id = " & .ListItems(InitFor)
            Conexao.Execute "DELETE from vendas_vendedores_Produto WHERE IDCliente = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Vendedores"
            Evento = "Excluir cliente"
            ID_documento = .ListItems(InitFor)
            Documento = "Vendedor: " & txtVendedor
            Documento1 = "Cliente: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cliente(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Cliente(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCliente_prod
    ProcCarregalistaClienteProd
    Frame13.Enabled = False
    Novo_vendedores3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procExcluirProduto_ProdCli()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista5
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)/serviço(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from vendas_vendedores_produto WHERE id = " & .ListItems(InitFor)
            '==================================
            Modulo = "Vendas/Vendedores"
            Evento = "Excluir produto/serviço"
            ID_documento = .ListItems(InitFor)
            Documento = "Vendedor: " & txtVendedor
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s)/serviço(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s)/serviço(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposCliente_prod2
    ProcCarregalistaClienteProd2
    Frame14.Enabled = False
    Novo_vendedores4 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_vendedores = True Or AlterarComissao = True Then
    USMsgBox ("Salve o vendedor antes de prosseguir."), vbExclamation, "CAPRIND v5.0"
    SSTab1.Tab = 0
    Permitido = False
    Exit Sub
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
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcValidarRegistros Lista, "Vendas/Vendedores"
    Case 9: ProcAjuda
    Case 10: ProcSair
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
    Case 5: ProcAjuda
    Case 6: ProcSair
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
    Case 1: ProcNovoCliente
    Case 2: ProcNovoProduto
    Case 3: procNovo_ProdCli
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
    Case 1: procSalvarCliente
    Case 2: ProcSalvarProduto
    Case 3: procSalvar_ProdCli
End Select

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
    Case 1: procExcluirCliente
    Case 2: ProcExcluirProduto
    Case 3: procExcluir_ProdCli
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcNovoMeta()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

FrameMeta.Enabled = True

txtDe.Text = ""
txtAte.Text = ""
txtMcomissao.Text = ""
txtIDMeta.Text = 0

txtDe.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarMeta()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtDe.Text = "" Then
    USMsgBox "Digite o valor inicial da meta!", vbInformation, "CAPRIND v5.0"
    txtDe.SetFocus
    Exit Sub
End If

If txtAte.Text = "" Then
    USMsgBox "Digite o valor Final da meta!", vbInformation, "CAPRIND v5.0"
    txtAte.SetFocus
    Exit Sub
End If

If txtMcomissao.Text = "" Then
    USMsgBox "Digite o valor da comissão para essa faixa de meta!", vbInformation, "CAPRIND v5.0"
    txtMcomissao.SetFocus
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_vendedores_Comissoes where id = " & txtIDMeta.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
Else
    If FunVerificaRegistroValidado("vendas_vendedores", "id = " & txtId, "vendedor", "o produto/serviço", "alterar", True, True) = False Then Exit Sub
End If
TBGravar!IDvendedor = txtId
TBGravar!De = txtDe.Text
TBGravar!Ate = txtAte.Text
TBGravar!Comissao = txtMcomissao.Text
TBGravar.Update
txtIDMeta = TBGravar!ID
TBGravar.Close
ProcCarregalistaMetas


    USMsgBox ("Dados gravados com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Gravar meta"
    
    If CodigoLista6 <> 0 And Lista6.ListItems.Count <> 0 Then
        Lista6.SelectedItem = Lista6.ListItems(CodigoLista6)
        Lista6.SetFocus
    End If

'==================================
'Modulo = "Vendas/Vendedores"
'ID_documento = txtIDMeta
'Documento = "Vendedor: " & txtVendedor
'Documento1 = "Meta"
'ProcGravaEvento
'==================================


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Vendedores.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Vendedores.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendedores.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendedores.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendedores.PageSize * (Pagina - 1)), 0), TBLISTA_Vendedores.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendedores.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLISTA_Vendedores!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendedores!n_vendedor), "", TBLISTA_Vendedores!n_vendedor)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Vendedores!vendedor), "", TBLISTA_Vendedores!vendedor)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendedores!Comissao), "", Format(TBLISTA_Vendedores!Comissao, "###,##0.00") & "%")
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendedores!DtValidacao), "Não", "Sim")
    End With
    TBLISTA_Vendedores.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Vendedores.RecordCount
If TBLISTA_Vendedores.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Vendedores.PageCount
ElseIf TBLISTA_Vendedores.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Vendedores.PageCount & " de: " & TBLISTA_Vendedores.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Vendedores.AbsolutePage - 1 & " de: " & TBLISTA_Vendedores.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoMeta
    Case 2: ProcSalvarMeta
    Case 3: procExcluirMeta
    Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

