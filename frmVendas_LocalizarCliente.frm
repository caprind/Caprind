VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_LocalizarCliente 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Administrativo | Vendas | Clientes - Localizar"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   820
      DibPicture      =   "frmVendas_LocalizarCliente.frx":0000
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_LocalizarCliente.frx":3650
   End
   Begin VB.Frame Frame9 
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
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   90
      TabIndex        =   26
      Top             =   7260
      Width           =   10785
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
         Left            =   5550
         TabIndex        =   14
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
         Left            =   2910
         TabIndex        =   13
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   7770
         TabIndex        =   18
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_LocalizarCliente.frx":396A
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
         Left            =   7230
         TabIndex        =   17
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_LocalizarCliente.frx":7111
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
         Left            =   6120
         TabIndex        =   15
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
         Left            =   6690
         TabIndex        =   16
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_LocalizarCliente.frx":AC20
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
         Left            =   8310
         TabIndex        =   19
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_LocalizarCliente.frx":ED12
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
         Left            =   8970
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de reg.: 0"
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
         TabIndex        =   28
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               reg. p/ pág."
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
         Left            =   2220
         TabIndex        =   27
         Top             =   240
         Width           =   2190
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   25
      Top             =   7890
      Width           =   10785
      _ExtentX        =   19024
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   4125
      Left            =   90
      TabIndex        =   12
      Top             =   3120
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   7276
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Cód."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "CNPJ/CPF"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Razão social"
         Object.Width           =   4736
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Endereço"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cidade"
         Object.Width           =   2663
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "E-mail"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "UF"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
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
      Height          =   1575
      Left            =   1500
      TabIndex        =   20
      Top             =   1530
      Width           =   9375
      Begin VB.Frame Frame4 
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
         Height          =   510
         Left            =   4380
         TabIndex        =   30
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
            TabIndex        =   10
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   11
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
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1110
         Width           =   8985
      End
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
         ItemData        =   "frmVendas_LocalizarCliente.frx":125A1
         Left            =   180
         List            =   "frmVendas_LocalizarCliente.frx":125BA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4125
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         ToolTipText     =   "Número do CNPJ."
         Top             =   1110
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbstatus 
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
         ItemData        =   "frmVendas_LocalizarCliente.frx":12610
         Left            =   180
         List            =   "frmVendas_LocalizarCliente.frx":1261D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Status."
         Top             =   1110
         Width           =   8985
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Número do CPF."
         Top             =   1110
         Visible         =   0   'False
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbFamilia 
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
         TabIndex        =   4
         ToolTipText     =   "Família."
         Top             =   1110
         Width           =   8985
      End
      Begin VB.Label Label9 
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
         Left            =   3937
         TabIndex        =   23
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label8 
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
         Left            =   1822
         TabIndex        =   22
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1575
      Left            =   90
      TabIndex        =   21
      Top             =   1530
      Width           =   1395
      Begin VB.OptionButton optFisica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Física"
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
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optJuridica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jurídica"
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
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   90
      TabIndex        =   24
      Top             =   540
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   1720
      ButtonCount     =   5
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
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
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
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   8700
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_LocalizarCliente.frx":1263F
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmVendas_LocalizarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia.Text <> "" Then
    txtTexto.Text = ""
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfiltrarpor = "Razão social" Or cmbfiltrarpor = "Nome fantasia" Or cmbfiltrarpor = "Cidade" Or cmbfiltrarpor = "Código do cliente" Then
    txtTexto.Visible = True
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = False
    cmbfamilia.Clear
    If cmbfiltrarpor = "Família" Then
        ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", False
    Else
        Set TBFamilia = CreateObject("adodb.recordset")
        TBFamilia.Open "Select * from Clientes_grupos where Texto <> 'Null' order by Texto", Conexao, adOpenKeyset, adLockOptimistic
        If TBFamilia.EOF = False Then
            Do While TBFamilia.EOF = False
                cmbfamilia.AddItem TBFamilia!Texto
                TBFamilia.MoveNext
            Loop
        End If
    End If
    TBFamilia.Close
End If
If cmbfiltrarpor = "Status" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = True
    txtcnpj.Visible = False
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optJuridica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = True
    txtCpf.Visible = False
End If
If cmbfiltrarpor = "CNPJ/CPF" And optFisica.Value = True Then
    txtTexto.Visible = False
    cmbfamilia.Visible = False
    cmbStatus.Visible = False
    txtcnpj.Visible = False
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbstatus_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbStatus.Text <> "" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If optFisica.Value = True Then
    TipoPessoa = "(C.tipo = 'FP' or C.tipo = 'FR')"
    TipoPessoaRel = "{C.tipo} = 'FP' or {C.tipo} = 'FR'"
    CpfCnpj = "C.cpf_cnpj = '" & txtCpf.Text & "'"
Else
    TipoPessoa = "(C.tipo = 'JP' or C.tipo = 'JR')"
    TipoPessoaRel = "({C.tipo} = 'JP' or {C.tipo} = 'JR')"
    CpfCnpj = "C.cpf_cnpj = '" & txtcnpj.Text & "'"
End If
If Telemarketing = False And Vendas_Proposta = False And Vendas_Analise = False Then ProspectoFiltro = "C.Prospecto = 'False'" Else ProspectoFiltro = "(C.Prospecto = 'False' or C.Prospecto = 'True')"
If Faturamento = True Then NFTexto = " and C.Enviar_NF = 'True'" Else NFTexto = ""

TextoFiltroVE = ""
INNERJOINTEXTOVE = ""
If Vendas_Proposta = True Or Vendas_PI = True Then
    With IIf(Vendas_Proposta = True, frmVendas_proposta, frmVendas_PI)
        If .txtVend_Ext <> "" Then
            INNERJOINTEXTOVE = " LEFT JOIN vendas_vendedores VV ON VV.N_Vendedor = " & .txtVE & " LEFT JOIN Vendas_Vendedores_Clientes VVC ON VVC.IDVendedor = VV.ID and VVC.IDCliente = C.IDCliente"
            TextoFiltroVE = " and (VV.Bloquear_venda_cliente = 'True' and VVC.IDCliente IS NOT NULL or VV.Bloquear_venda_cliente = 'False')"
        End If
    End With
End If

CamposFiltro = "C.idTipoEmpresa, C.IDCliente, C.CPF_CNPJ, C.NomeRazao, C.Tipo_endereco, C.Endereco, C.Cidade, C.Email, C.UF, C.Tipo"
INNERJOINTEXTO = "Select " & CamposFiltro & " from clientes C LEFT JOIN vendas_tele VT ON C.idcliente = VT.IDCliente LEFT JOIN compras_fornecedores_familia CFF ON C.IDCliente = CFF.IDCliente LEFT JOIN Clientes_grupos CG ON C.IDGrupo = CG.ID" & INNERJOINTEXTOVE
TextoFiltroPadrao = TipoPessoa & " and " & ProspectoFiltro & " and C.DtValidacao IS NOT NULL and C.status <> 'Bloqueado'" & TextoFiltroVE & NFTexto & " group by " & CamposFiltro & " order by C.nomerazao"

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Or cmbStatus.Visible = True And cmbStatus <> "" Or txtcnpj.Visible = True And txtcnpj <> "__.___.___/____-__" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Then
    If cmbfiltrarpor = "Status" Then
        If cmbStatus.Text = "Bloqueado" Then TextoFiltro = "VT.bloqueado = 'True'" Else TextoFiltro = "C.status = '" & cmbStatus.Text & "'"
        StrSqlLocCliPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "Família" Then
            StrSqlLocCliPadrao = INNERJOINTEXTO & " where CFF.Familia = '" & cmbfamilia & "' and CFF.tipo = 'C' and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Grupo" Then
                StrSqlLocCliPadrao = INNERJOINTEXTO & " where CG.Texto = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "CNPJ/CPF" Then
                    StrSqlLocCliPadrao = INNERJOINTEXTO & " where " & CpfCnpj & " and " & TextoFiltroPadrao
                ElseIf cmbfiltrarpor = "Código do cliente" Then
                        StrSqlLocCliPadrao = INNERJOINTEXTO & " where C.IDCliente = " & txtTexto & " and " & TextoFiltroPadrao
                    Else
                        Select Case cmbfiltrarpor
                            Case "Razão social": TextoFiltro = "C.nomerazao"
                            Case "Nome fantasia": TextoFiltro = "C.nomefantasia"
                            Case "Cidade": TextoFiltro = "C.cidade"
                        End Select
                        StrSqlLocCliPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSqlLocCliPadrao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

'Debug.print StrSqlLocCliPadrao

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_cliente_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_cliente_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_cliente_padrao.PageCount - 1)
    Else
        TBLocalizar_cliente_padrao.AbsolutePage = TBLocalizar_cliente_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)
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
    TBLocalizar_cliente_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_cliente_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_cliente_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_cliente_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_cliente_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_cliente_padrao.AbsolutePage = TBLocalizar_cliente_padrao.PageCount
ProcExibePagina (TBLocalizar_cliente_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10785, 4, True

If Estoque_Consignacao = True Then Caption = "Estoque - Recebimento - Consignação - Localizar cliente"
If Estoque_Inventario = True Then Caption = "Estoque - Inventário - Localizar cliente"
If Vendas_Proposta = True Then Caption = "Administrativo - Vendas - Proposta comercial - Localizar cliente"
If Vendas_PI = True Then Caption = "Administrativo - Vendas - Pedido interno - Localizar cliente"
If Financeiro_Contas_Pagar = True Then Caption = "Administrativo - Financeiro - Contas à pagar - Localizar cliente"
If Financeiro_Contas_Pagas = True Then Caption = "Administrativo - Financeiro - Contas pagas - Localizar cliente"
If Financeiro_Contas_Receber = True Then Caption = "Administrativo - Financeiro - Contas à receber - Localizar cliente"
If Financeiro_Contas_Recebidas = True Then Caption = "Administrativo - Financeiro - Contas recebidas - Localizar cliente"
If Faturamento = True Then
    If Sit_REG = 4 Then
        Caption = "Administrativo - Faturamento - Minuta de despacho - Localizar cliente"
    Else
        If Formulario = "Faturamento/Nota fiscal/Própria" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Localizar cliente"
        ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
                Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Localizar cliente"
            ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                    Caption = "Estoque - Ordem de faturamento - Localizar cliente"
                Else
                    Caption = "Estoque - Nota fiscal - Localizar cliente"
        End If
    End If
End If
If Engenharia_Produtos = True Then TextoCaption = "Engenharia"
If Compras_Produtos = True Then TextoCaption = "Compras"
If Vendas_Produtos = True Then TextoCaption = "Vendas"
If Engenharia_Localcliente = True Then Caption = TextoCaption & " - Produtos e serviços - Cadastro de códigos de referência - Localizar cliente"
If Engenharia_Localcliente1 = True Then Caption = TextoCaption & " - Produtos e serviços - Localizar cliente"
If PCP_Ordem = True Then Caption = "PCP - Gerenciamento de ordem - Localizar cliente"
If Vendas_Analise = True Then Caption = "Outros - Análise crítica - Localizar cliente"
If Vendas_Programacao = True Then Caption = "Administrativo - Vendas - Programação - Localizar cliente"
If Compras_Pedido = True Then Caption = "Administrativo - Compras - Pedido - Localizar cliente"
If Clientes = True Then Caption = "Administrativo - Vendas - Clientes - Localizar fornecedor"
If Vendas_Vendedores = True Then Caption = "Administrativo - Vendas - Vendedores - Localizar cliente"
If Qualidade_PPAP_PSW = True Then Caption = "Qualidade - PPAP - PSW - Localizar cliente"
If Qualidade_PPAP_FMEA = True Then Caption = "Qualidade - PPAP - FMEA - Localizar cliente"
If Compras_Fornecedores = True Then Caption = "Compras - Fornecedores - Localizar cliente"
If Compras_Fornecedores = True Then Caption = "Compras - Fornecedores - Localizar cliente"
If Fiscal_NaturezaOperacao = True Then Caption = "Faturamento - Fiscal - Natureza de operação"
cmbfiltrarpor = "Nome fantasia"

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

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optFisica_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If optFisica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = False
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optJuridica_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If optJuridica.Value = True And cmbfiltrarpor = "CNPJ/CPF" Then
    txtTexto.Visible = False
    txtTexto = ""
    cmbfamilia.Visible = False
    cmbfamilia.ListIndex = -1
    cmbStatus.Visible = False
    cmbStatus.ListIndex = -1
    txtcnpj.Visible = True
    txtCpf.Visible = False
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

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
If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
    txtCpf.Text = "___.___.___-__"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcnpj_Change()
On Error GoTo tratar_erro
  
ListView1.ListItems.Clear
If txtcnpj.Text <> "__.___.___/____-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtCpf.Text = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCpf_Change()
On Error GoTo tratar_erro
  
ListView1.ListItems.Clear
If txtCpf.Text <> "___.___.___-__" Then
    txtTexto.Text = ""
    cmbfamilia.ListIndex = -1
    cmbStatus.ListIndex = -1
    txtcnpj.Text = "__.___.___/____-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
    Case vbKeyReturn: ListView1_DblClick
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * FROM Clientes WHERE idcliente = " & ListView1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    If Estoque_Consignacao = True Then
        With frmEstoque_Recebimento_consignacao
            .txtCliente.Text = ListView1.SelectedItem.SubItems(2)
            .txtid_cliente.Text = ListView1.SelectedItem
            .Txt_tipodest = "C"
        End With
    ElseIf Estoque_Inventario = True Then
            With frmestoque_fisico
                .Cmb_tipo_cli_forn = "Cliente"
                .Txt_ID_cli_forn = ListView1.SelectedItem
                .Txt_cli_forn = ListView1.SelectedItem.SubItems(2)
            End With
        ElseIf Vendas_Programacao = True Then
                With frmVendas_programacao
                    If Left(TBClientes!Tipo, 1) = "J" And TBClientes!idTipoEmpresa = 1 Then
                        If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
                    End If
                    .txtID_cli = ListView1.SelectedItem
                    .txtCliente = ListView1.SelectedItem.ListSubItems(2)
                End With
            ElseIf Vendas_Analise = True Then
                    With frmVendas_analise
                        .txtIDcliente = ListView1.SelectedItem
                        .txtCliente = ListView1.SelectedItem.ListSubItems(2)
                    End With
                ElseIf Vendas_Proposta = True Or Vendas_PI = True Then
                        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
                            If Sit_REG = 1 Then
                                If Left(TBClientes!Tipo, 1) = "J" And TBClientes!idTipoEmpresa = 1 Then
                                    If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
                                End If
                                .txtIDcliente = ""
                                .txtIDcliente.Text = ListView1.SelectedItem
                            Else
                                .cmbtransportadora = ListView1.SelectedItem.ListSubItems(2)
                            End If
                        End With
                    ElseIf Financeiro_Contas_Pagar = True Then
                            With frmContas_Pagar
                                .Cmb_tipo = "Cliente"
                                .txtIDFornec = ListView1.SelectedItem
                            End With
                        ElseIf Financeiro_Contas_Pagas = True Then
                                With frmContas_Pagas
                                    .Cmb_tipo = "Cliente"
                                    .txtIDFornec = ListView1.SelectedItem
                                End With
                            ElseIf Financeiro_Contas_Receber = True Then
                                    With frmContas_Receber
                                        .Cmb_tipo = "Cliente"
                                        .txtIDcliente = ListView1.SelectedItem
                                        Unload Me
                                        Exit Sub
                                    End With
                                ElseIf Financeiro_Contas_Recebidas = True Then
                                        With frmContas_recebidas
                                            .Cmb_tipo = "Cliente"
                                            .txtIDcliente = ListView1.SelectedItem
                                        End With
                                    ElseIf Engenharia_Localcliente = True Then
                                            With frmproj_produto_referencia
                                                .Txt_ID_cliente_forn = ListView1.SelectedItem
                                                .Txt_tipo = "C"
                                                .txtAplicacao = ListView1.SelectedItem.ListSubItems(2)
                                            End With
                                        ElseIf Engenharia_Localcliente1 = True Then
                                                With frmproj_produto
                                                    .cmbcliente.Clear
                                                    .txtIDcliente = ListView1.SelectedItem
                                                    .cmbcliente.DataField = .txtIDcliente
                                                    .cmbcliente.AddItem ListView1.SelectedItem.ListSubItems(2)
                                                    .cmbcliente = ListView1.SelectedItem.ListSubItems(2)
                                                    .txtRevenda_forn = "0,00000"
                                                    .txtConsumo_forn = "0,00000"
                                                End With
                                            ElseIf PCP_Ordem = True Then
                                                    With frmprod
                                                        .Txt_ID_cliente = ListView1.SelectedItem
                                                        .txtCliente = ListView1.SelectedItem.ListSubItems(2)
                                                    End With
                                                ElseIf RNC = True Then
                                                        With frmQualidade_RNC
                                                            .txtID_forn = ListView1.SelectedItem
                                                            .txtFornecedor = ListView1.SelectedItem.ListSubItems(2)
                                                            .txttipo = "C"
                                                        End With
                                                    ElseIf Engenharia_Normas = True Then
                                                            With frmNorma
                                                                .Txt_ID_cliente = ListView1.SelectedItem
                                                                .Txt_cliente = ListView1.SelectedItem.ListSubItems(2)
                                                            End With
                                                        ElseIf Compras_Pedido = True Then
                                                                frmCompras_Pedido.cmbtransporte = ListView1.SelectedItem.ListSubItems(2)
                                                            ElseIf Clientes = True Then
                                                                    frmVendas_cliente.cmbtransportadora = ListView1.SelectedItem.ListSubItems(2)
                                                                ElseIf Vendas_Vendedores = True Then
                                                                        With frmVendas_Vendedores
                                                                            If Aplic = 1 Then
                                                                                .txtIDcliente = ListView1.SelectedItem
                                                                                .txtnomerazao = ListView1.SelectedItem.ListSubItems(2)
                                                                                .txtCidade = ListView1.SelectedItem.ListSubItems(4)
                                                                            Else
                                                                                .txtIDCliente_prod = ListView1.SelectedItem
                                                                                .txtCliente_prod = ListView1.SelectedItem.ListSubItems(2)
                                                                                .txtCidadeCliente_prod = ListView1.SelectedItem.ListSubItems(4)
                                                                            End If
                                                                        End With
                                                                    ElseIf Qualidade_PPAP_PSW = True Then
                                                                            With frmQualidadePPAP
                                                                                .txtIDcliente = ListView1.SelectedItem
                                                                                .txtCliente = ListView1.SelectedItem.ListSubItems(2)
                                                                            End With
                                                                        ElseIf Qualidade_PPAP_FMEA = True Then
                                                                                With frmQualidadePPAP_FMEA
                                                                                    .txtIDcliente = ListView1.SelectedItem
                                                                                    .txtCliente = ListView1.SelectedItem.ListSubItems(2)
                                                                                End With
                                                                            ElseIf Compras_Fornecedores = True Then
                                                                                    frmCompras_fornecedores.cmbtransportadora = ListView1.SelectedItem.ListSubItems(2)
                                                                                ElseIf Fiscal_NaturezaOperacao = True Then
                                                                                        frm_Natureza_OP.txtid_cliente = ListView1.SelectedItem
                                                                                    ElseIf Faturamento = True Then
                                                                                            If Sit_REG < 4 Then
                                                                                             If Formulario <> "Estoque/Ordem de faturamento" Then
                                                                                                With frmFaturamento_Prod_Serv
                                                                                                    If Sit_REG = 1 Then
                                                                                                        If Left(TBClientes!Tipo, 1) = "J" And TBClientes!idTipoEmpresa = 1 Then
                                                                                                            'If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
                                                                                                        End If
                                                                                                                    
                                                                                                        IDCliente = ListView1.SelectedItem
                                                                                                        .txt_Razao.Text = ListView1.SelectedItem.ListSubItems(2)
                                                                                                        If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                            Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                        Else
                                                                                                            Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                        End If
                                                                                                        .txt_Endereco.Text = Endereco
                                                                                                        .txtNumero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
                                                                                                        If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                                                                                                            Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
                                                                                                        Else
                                                                                                            Bairro = IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
                                                                                                        End If
                                                                                                        .txt_Bairro.Text = Bairro
                                                                                                        .txttipocliente.Text = IIf(IsNull(TBClientes!Tipo), "", TBClientes!Tipo)
                                                                                                        If TBClientes!idTipoEmpresa = 1 Then .txt_CNPJ_CPF.Text = IIf(IsNull(TBClientes!CPF_CNPJ), "", TBClientes!CPF_CNPJ)
                                                                                                        .Txt_CEP = IIf(IsNull(TBClientes!CEP), "", TBClientes!CEP)
                                                                                                        If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then .txt_IE.Text = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE) Else .txt_IE = IIf(IsNull(TBClientes!RG_IM), "", TBClientes!RG_IM)
                                                                                                        .txt_Municipio.Text = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
                                                                                                        .cbo_UF.Text = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
                                                                                                        .txt_FoneFAX.Text = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
                                                                                                        If TBClientes!chkSuframa = True Then Suframa = True Else Suframa = False
                                                                                                        .txtIDcliente.Text = IDCliente
                                                                                                    ElseIf Sit_REG = 2 Then
                                                                                                            .txtidinttransp = ListView1.SelectedItem
                                                                                                            .TxtTransp_nome.Text = ListView1.SelectedItem.ListSubItems(2)
                                                                                                            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            Else
                                                                                                                Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            End If
                                                                                                            .txtTransp_endereco = Endereco
                                                                                                            .txtTransp_numero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
                                                                                                            .txtTransp_municipio = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
                                                                                                            .txtTransp_uf_Transportadora = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
                                                                                                            If TBClientes!idTipoEmpresa = 1 Then
                                                                                                                If IsNull(TBClientes!CPF_CNPJ) = True Or TBClientes!CPF_CNPJ = "__.___.___/____-__" Or TBClientes!CPF_CNPJ = "" Then .txtTransp_cnpj = "" Else .txtTransp_cnpj = TBClientes!CPF_CNPJ
                                                                                                            End If
                                                                                                            If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then
                                                                                                                .txtTransp_IE = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
                                                                                                                .txtTransp_IE.Locked = False
                                                                                                                .txtTransp_IE.TabStop = True
                                                                                                            Else
                                                                                                                .txtTransp_IE.Locked = True
                                                                                                                .txtTransp_IE.TabStop = False
                                                                                                            End If
                                                                                                        Else
                                                                                                            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            Else
                                                                                                                Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            End If
                                                                                                            Redespacho = "Nome: " & ListView1.SelectedItem.ListSubItems(2) & " - Endereço: " & Endereco & " - Número: " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - Cidade: " & IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade) & " - UF: " & IIf(IsNull(TBClientes!UF), "", TBClientes!UF) & " - CNPJ: " & IIf(TBClientes!idTipoEmpresa = 1, IIf(IsNull(TBClientes!CPF_CNPJ), "", TBClientes!CPF_CNPJ), "") & " - IE: " & IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
                                                                                                            If .txtDados_DadosAdicionais <> "" Then
                                                                                                                .txtDados_DadosAdicionais = .txtDados_DadosAdicionais & " | REDESPACHO: " & Redespacho
                                                                                                            Else
                                                                                                                .txtDados_DadosAdicionais = Redespacho
                                                                                                            End If
                                                                                                    End If
                                                                                                End With
                                                                                                Else
                                                                                                With frmEstoque_Ordem_Faturamento
                                                                                                    If Sit_REG = 1 Then
                                                                                                        If Left(TBClientes!Tipo, 1) = "J" And TBClientes!idTipoEmpresa = 1 Then
                                                                                                            'If FunVerifRegimeTribCliForn(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), False, True) = False Then Exit Sub
                                                                                                        End If
                                                                                                                    
                                                                                                        IDCliente = ListView1.SelectedItem
                                                                                                        .txt_Razao.Text = ListView1.SelectedItem.ListSubItems(2)
                                                                                                        If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                            Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                        Else
                                                                                                            Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                        End If
                                                                                                        .txt_Endereco.Text = Endereco
                                                                                                        .txtNumero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
                                                                                                        If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                                                                                                            Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
                                                                                                        Else
                                                                                                            Bairro = IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
                                                                                                        End If
                                                                                                        .txt_Bairro.Text = Bairro
                                                                                                        .txttipocliente.Text = IIf(IsNull(TBClientes!Tipo), "", TBClientes!Tipo)
                                                                                                        If TBClientes!idTipoEmpresa = 1 Then .txt_CNPJ_CPF.Text = IIf(IsNull(TBClientes!CPF_CNPJ), "", TBClientes!CPF_CNPJ)
                                                                                                        .Txt_CEP = IIf(IsNull(TBClientes!CEP), "", TBClientes!CEP)
                                                                                                        If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then .txt_IE.Text = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE) Else .txt_IE = IIf(IsNull(TBClientes!RG_IM), "", TBClientes!RG_IM)
                                                                                                        .txt_Municipio.Text = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
                                                                                                        .cbo_UF.Text = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
                                                                                                        .txt_FoneFAX.Text = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
                                                                                                        If TBClientes!chkSuframa = True Then Suframa = True Else Suframa = False
                                                                                                        .txtIDcliente.Text = IDCliente
                                                                                                    ElseIf Sit_REG = 2 Then
                                                                                                            .txtidinttransp = ListView1.SelectedItem
                                                                                                            .TxtTransp_nome.Text = ListView1.SelectedItem.ListSubItems(2)
                                                                                                            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            Else
                                                                                                                Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            End If
                                                                                                            .txtTransp_endereco = Endereco
                                                                                                            .txtTransp_numero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
                                                                                                            .txtTransp_municipio = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
                                                                                                            .txtTransp_uf_Transportadora = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
                                                                                                            If TBClientes!idTipoEmpresa = 1 Then
                                                                                                                If IsNull(TBClientes!CPF_CNPJ) = True Or TBClientes!CPF_CNPJ = "__.___.___/____-__" Or TBClientes!CPF_CNPJ = "" Then .txtTransp_cnpj = "" Else .txtTransp_cnpj = TBClientes!CPF_CNPJ
                                                                                                            End If
                                                                                                            If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then
                                                                                                                .txtTransp_IE = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
                                                                                                                .txtTransp_IE.Locked = False
                                                                                                                .txtTransp_IE.TabStop = True
                                                                                                            Else
                                                                                                                .txtTransp_IE.Locked = True
                                                                                                                .txtTransp_IE.TabStop = False
                                                                                                            End If
                                                                                                        Else
                                                                                                            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            Else
                                                                                                                Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                            End If
                                                                                                            Redespacho = "Nome: " & ListView1.SelectedItem.ListSubItems(2) & " - Endereço: " & Endereco & " - Número: " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - Cidade: " & IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade) & " - UF: " & IIf(IsNull(TBClientes!UF), "", TBClientes!UF) & " - CNPJ: " & IIf(TBClientes!idTipoEmpresa = 1, IIf(IsNull(TBClientes!CPF_CNPJ), "", TBClientes!CPF_CNPJ), "") & " - IE: " & IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
                                                                                                    End If
                                                                                                End With
                                                                                                
                                                                                                End If
                                                                                            Else
                                                                                                With frmMinuta
                                                                                                    .txtID_transp = ListView1.SelectedItem
                                                                                                    .txtTranportadora = ListView1.SelectedItem.ListSubItems(2)
                                                                                                    If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                                                                                                        Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                    Else
                                                                                                        Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
                                                                                                    End If
                                                                                                    .txtendereco = Endereco
                                                                                                    .txtCidade = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
                                                                                                    .cmbuf = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
                                                                                                    .txttelefone = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
                                                                                                    .txtFax = IIf(IsNull(TBClientes!Fax), "", TBClientes!Fax)
                                                                                                    If TBClientes!idTipoEmpresa = 1 Then
                                                                                                        If IsNull(TBClientes!CPF_CNPJ) = True Or TBClientes!CPF_CNPJ = "__.___.___/____-__" Or TBClientes!CPF_CNPJ = "" Then .txtcnpj = "" Else .txtcnpj = TBClientes!CPF_CNPJ
                                                                                                    End If
                                                                                                    .txtIE = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
                                                                                                End With
                                                                                            End If
        End If
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocCliPadrao = "" Then Exit Sub
Set TBLocalizar_cliente_padrao = CreateObject("adodb.recordset")
TBLocalizar_cliente_padrao.Open StrSqlLocCliPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_cliente_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_cliente_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_cliente_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_cliente_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_cliente_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_cliente_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_cliente_padrao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLocalizar_cliente_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_cliente_padrao!IDCliente
        If TBLocalizar_cliente_padrao!idTipoEmpresa = 1 And IsNull(TBLocalizar_cliente_padrao!CPF_CNPJ) = False And TBLocalizar_cliente_padrao!CPF_CNPJ <> "__.___.___/____-__" Then .Item(.Count).SubItems(1) = TBLocalizar_cliente_padrao!CPF_CNPJ
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_cliente_padrao!NomeRazao), "", TBLocalizar_cliente_padrao!NomeRazao)
        If IsNull(TBLocalizar_cliente_padrao!Tipo_endereco) = False And TBLocalizar_cliente_padrao!Tipo_endereco <> "" Then
            Endereco = TBLocalizar_cliente_padrao!Tipo_endereco & ": " & IIf(IsNull(TBLocalizar_cliente_padrao!Endereco), "", TBLocalizar_cliente_padrao!Endereco)
        Else
            Endereco = IIf(IsNull(TBLocalizar_cliente_padrao!Endereco), "", TBLocalizar_cliente_padrao!Endereco)
        End If
        .Item(.Count).SubItems(3) = Endereco
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLocalizar_cliente_padrao!Cidade), "", TBLocalizar_cliente_padrao!Cidade)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLocalizar_cliente_padrao!Email), "", TBLocalizar_cliente_padrao!Email)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLocalizar_cliente_padrao!UF), "", TBLocalizar_cliente_padrao!UF)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLocalizar_cliente_padrao!Tipo), "", TBLocalizar_cliente_padrao!Tipo)
    End With
    TBLocalizar_cliente_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLocalizar_cliente_padrao.RecordCount
If TBLocalizar_cliente_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_cliente_padrao.PageCount
ElseIf TBLocalizar_cliente_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_cliente_padrao.PageCount & " de: " & TBLocalizar_cliente_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_cliente_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_cliente_padrao.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
