VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContas_Pagar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Contas a pagar"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmContas_pagar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
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
   Begin VB.TextBox txtidintconta 
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
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Número da conta."
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   60
      TabIndex        =   72
      Top             =   8610
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
         ItemData        =   "frmContas_pagar.frx":1042
         Left            =   6960
         List            =   "frmContas_pagar.frx":1055
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   187
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
         TabIndex        =   37
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
         TabIndex        =   39
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   11760
         TabIndex        =   43
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_pagar.frx":1086
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
         TabIndex        =   42
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_pagar.frx":482A
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
         TabIndex        =   40
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
         TabIndex        =   41
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_pagar.frx":8333
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
         TabIndex        =   44
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmContas_pagar.frx":C422
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
         Left            =   5610
         TabIndex        =   84
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label18 
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
         Index           =   0
         Left            =   2040
         TabIndex        =   81
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
         TabIndex        =   74
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
         TabIndex        =   73
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3435
      Left            =   55
      TabIndex        =   53
      Top             =   990
      Width           =   15195
      Begin VB.CheckBox chkConta_fixa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Conta fixa"
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
         Left            =   180
         TabIndex        =   33
         Top             =   3000
         Width           =   1365
      End
      Begin VB.CommandButton Cmd_localizar_contatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14700
         Picture         =   "frmContas_pagar.frx":FCAE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Localizar contatos."
         Top             =   975
         Width           =   315
      End
      Begin VB.CheckBox Chk_agendado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agendado"
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
         Left            =   180
         TabIndex        =   32
         Top             =   2745
         Width           =   1365
      End
      Begin VB.CheckBox Chk_devolucao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Devolução"
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
         Left            =   180
         TabIndex        =   31
         Top             =   2505
         Width           =   1365
      End
      Begin VB.CheckBox Chk_antecipacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Antecipação"
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
         Left            =   180
         TabIndex        =   30
         Top             =   2250
         Width           =   1365
      End
      Begin VB.TextBox txt_Competencia 
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
         Left            =   8310
         MaxLength       =   30
         TabIndex        =   26
         ToolTipText     =   "Competência."
         Top             =   1560
         Width           =   1320
      End
      Begin VB.CommandButton Cmd_competencia 
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
         Left            =   9615
         Picture         =   "frmContas_pagar.frx":FFC2
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Filtrar por número do documento."
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton Cmd_valor 
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
         Left            =   5190
         Picture         =   "frmContas_pagar.frx":103DD
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Filtrar por valor."
         Top             =   975
         Width           =   315
      End
      Begin VB.CommandButton Cmd_data_transacao 
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
         Left            =   7980
         Picture         =   "frmContas_pagar.frx":107F8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Filtrar por data da transação."
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton Cmd_localizar_tipo_dcto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9780
         Picture         =   "frmContas_pagar.frx":10C13
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Localizar tipo do documento."
         Top             =   370
         Width           =   315
      End
      Begin VB.ComboBox Cmb_tipo 
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
         ItemData        =   "frmContas_pagar.frx":10D15
         Left            =   5580
         List            =   "frmContas_pagar.frx":10D25
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Tipo."
         Top             =   975
         Width           =   1905
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
         ItemData        =   "frmContas_pagar.frx":10D61
         Left            =   180
         List            =   "frmContas_pagar.frx":10D63
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   370
         Width           =   6570
      End
      Begin MSMask.MaskEdBox txtParcela 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         ToolTipText     =   "Número da parcela."
         Top             =   975
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdForma 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7920
         Picture         =   "frmContas_pagar.frx":10D65
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Localizar forma da baixa."
         Top             =   1560
         Width           =   315
      End
      Begin VB.ComboBox cmb_forma 
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
         ItemData        =   "frmContas_pagar.frx":10E67
         Left            =   4440
         List            =   "frmContas_pagar.frx":10E69
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Forma da baixa prevista."
         Top             =   1560
         Width           =   3465
      End
      Begin VB.TextBox txtValorTotal 
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
         Left            =   4125
         TabIndex        =   15
         ToolTipText     =   "Valor."
         Top             =   975
         Width           =   1050
      End
      Begin VB.ComboBox txtNPedido 
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
         Left            =   13425
         TabIndex        =   8
         ToolTipText     =   "Número do pedido de compra."
         Top             =   370
         Width           =   1280
      End
      Begin VB.ComboBox cmbBanco 
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
         ItemData        =   "frmContas_pagar.frx":10E6B
         Left            =   180
         List            =   "frmContas_pagar.frx":10E6D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Instituição bancária prevista."
         Top             =   1560
         Width           =   4245
      End
      Begin VB.CommandButton Cmdlocalizarforn 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14370
         Picture         =   "frmContas_pagar.frx":10E6F
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Localizar fornecedor."
         Top             =   975
         Width           =   315
      End
      Begin VB.TextBox txtstatus 
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
         Left            =   10005
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   1560
         Width           =   4680
      End
      Begin VB.TextBox txtIDFornec 
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
         Left            =   7500
         MaxLength       =   20
         TabIndex        =   18
         ToolTipText     =   "Código do fornecedor."
         Top             =   975
         Width           =   810
      End
      Begin VB.CommandButton cmdstatus 
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
         Left            =   14700
         Picture         =   "frmContas_pagar.frx":10F71
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Filtrar por status."
         Top             =   1560
         Width           =   315
      End
      Begin VB.ComboBox cmbtipo_conta 
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
         ItemData        =   "frmContas_pagar.frx":1138C
         Left            =   8370
         List            =   "frmContas_pagar.frx":1138E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Tipo do documento."
         Top             =   370
         Width           =   1065
      End
      Begin VB.CommandButton cmdtipo 
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
         Left            =   9450
         Picture         =   "frmContas_pagar.frx":11390
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Filtrar por tipo do documento."
         Top             =   370
         Width           =   315
      End
      Begin VB.CommandButton cmdpedido 
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
         Left            =   14700
         Picture         =   "frmContas_pagar.frx":117AB
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Filtrar por número do pedido de compra."
         Top             =   370
         Width           =   315
      End
      Begin VB.CommandButton cmdemissao 
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
         Left            =   1380
         Picture         =   "frmContas_pagar.frx":11BC6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Filtrar por data de emissão."
         Top             =   975
         Width           =   315
      End
      Begin VB.CommandButton cmdvencimento 
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
         Left            =   2970
         Picture         =   "frmContas_pagar.frx":11FE1
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Filtrar por data de vencimento."
         Top             =   975
         Width           =   315
      End
      Begin VB.CommandButton cmddoc 
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
         Left            =   13035
         Picture         =   "frmContas_pagar.frx":123FC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Filtrar por número do documento."
         Top             =   370
         Width           =   315
      End
      Begin VB.TextBox txtNDocumento 
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
         Left            =   10170
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Número do documento."
         Top             =   370
         Width           =   2850
      End
      Begin VB.CommandButton cmdLocalizar_fornecedor 
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
         Left            =   14025
         Picture         =   "frmContas_pagar.frx":12817
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Filtrar por fornecedor."
         Top             =   975
         Width           =   315
      End
      Begin MSComCtl2.DTPicker txtDtEmissao 
         Height          =   315
         Left            =   180
         TabIndex        =   10
         ToolTipText     =   "Data de emissão."
         Top             =   975
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   199426051
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker txtDtpagto 
         Height          =   315
         Left            =   1770
         TabIndex        =   12
         ToolTipText     =   "Data de vencimento."
         Top             =   975
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   199426051
         CurrentDate     =   39057
      End
      Begin VB.TextBox txtFornec 
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
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Nome do fornecedor."
         Top             =   975
         Width           =   5700
      End
      Begin MSComctlLib.ListView Lista_PC 
         Height          =   1095
         Left            =   7515
         TabIndex        =   35
         Top             =   2190
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   1931
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   6544
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtobs 
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
         Height          =   1095
         Left            =   1650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         ToolTipText     =   "Observações."
         Top             =   2190
         Width           =   5820
      End
      Begin MSComCtl2.DTPicker Txt_data_transacao 
         Height          =   315
         Left            =   6772
         TabIndex        =   1
         ToolTipText     =   "Data da transação."
         Top             =   370
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   171638787
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Instituição bancária prevista"
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
         Left            =   1155
         TabIndex        =   85
         Top             =   1350
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Competência"
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
         Left            =   8505
         TabIndex        =   83
         Top             =   1350
         Width           =   930
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. transação"
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
         Left            =   6877
         TabIndex        =   77
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contas contábeis"
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
         Left            =   10530
         TabIndex        =   76
         Top             =   1980
         Width           =   1470
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   6382
         TabIndex        =   75
         Top             =   780
         Width           =   300
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa*"
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
         Left            =   3045
         TabIndex        =   69
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma da baixa prevista"
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
         Left            =   5302
         TabIndex        =   67
         Top             =   1350
         Width           =   1740
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela*"
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
         Left            =   3420
         TabIndex        =   63
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   12068
         TabIndex        =   62
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo docto.*"
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
         Left            =   8445
         TabIndex        =   61
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor*"
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
         Left            =   10718
         TabIndex        =   60
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. vencto."
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
         Left            =   1943
         TabIndex        =   59
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ped. compra"
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
         Left            =   13608
         TabIndex        =   58
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor*"
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
         Left            =   4365
         TabIndex        =   57
         Top             =   780
         Width           =   570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
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
         Left            =   4035
         TabIndex        =   56
         Top             =   1980
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. emissão"
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
         Left            =   360
         TabIndex        =   55
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº documento*"
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
         Left            =   11040
         TabIndex        =   54
         Top             =   180
         Width           =   1110
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   71
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   18
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
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
      ButtonLeft3     =   81
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
      ButtonLeft4     =   127
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
      ButtonLeft5     =   174
      ButtonTop5      =   2
      ButtonWidth5    =   60
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "C. contábil"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Conta contábil (F6)"
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
      ButtonLeft6     =   236
      ButtonTop6      =   2
      ButtonWidth6    =   66
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Agenda"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Agenda do dia (F7)"
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
      ButtonLeft7     =   304
      ButtonTop7      =   2
      ButtonWidth7    =   51
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Parcelar"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Parcelar (F8)"
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
      ButtonLeft8     =   357
      ButtonTop8      =   2
      ButtonWidth8    =   55
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Copiar"
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Copiar (F9)"
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
      ButtonLeft9     =   414
      ButtonTop9      =   2
      ButtonWidth9    =   44
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Baixar"
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Baixar (F10)"
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
      ButtonLeft10    =   460
      ButtonTop10     =   2
      ButtonWidth10   =   44
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Status"
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Status (F11)"
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
      ButtonLeft11    =   506
      ButtonTop11     =   2
      ButtonWidth11   =   45
      ButtonHeight11  =   21
      ButtonUseMaskColor11=   0   'False
      ButtonCaption12 =   "Centro de custo"
      ButtonEnabled12 =   0   'False
      ButtonIconSize12=   32
      ButtonToolTipText12=   "Centro de custo (F12)"
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
      ButtonLeft12    =   553
      ButtonTop12     =   2
      ButtonWidth12   =   97
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
      ButtonCaption13 =   "Visualizar"
      ButtonEnabled13 =   0   'False
      ButtonIconSize13=   32
      ButtonToolTipText13=   "Visualizar contas relacionadas."
      ButtonKey13     =   "13"
      ButtonAlignment13=   2
      BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState13   =   5
      ButtonLeft13    =   652
      ButtonTop13     =   2
      ButtonWidth13   =   52
      ButtonHeight13  =   21
      ButtonUseMaskColor13=   0   'False
      ButtonCaption14 =   "Agendar"
      ButtonEnabled14 =   0   'False
      ButtonIconSize14=   32
      ButtonToolTipText14=   "Agendar pagamento."
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
      ButtonLeft14    =   706
      ButtonTop14     =   2
      ButtonWidth14   =   56
      ButtonHeight14  =   21
      ButtonUseMaskColor14=   0   'False
      ButtonEnabled15 =   0   'False
      ButtonIconSize15=   32
      ButtonAlignment15=   2
      ButtonType15    =   1
      ButtonStyle15   =   -1
      BeginProperty ButtonFont15 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState15   =   -1
      ButtonLeft15    =   764
      ButtonTop15     =   4
      ButtonWidth15   =   2
      ButtonHeight15  =   54
      ButtonCaption16 =   "Ajuda"
      ButtonEnabled16 =   0   'False
      ButtonIconSize16=   32
      ButtonToolTipText16=   "Ajuda (F1)"
      ButtonKey16     =   "16"
      ButtonAlignment16=   2
      BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft16    =   768
      ButtonTop16     =   2
      ButtonWidth16   =   41
      ButtonHeight16  =   21
      ButtonUseMaskColor16=   0   'False
      ButtonCaption17 =   "Sair"
      ButtonEnabled17 =   0   'False
      ButtonIconSize17=   32
      ButtonToolTipText17=   "Sair (Esc)"
      ButtonKey17     =   "17"
      ButtonAlignment17=   2
      BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft17    =   811
      ButtonTop17     =   2
      ButtonWidth17   =   30
      ButtonHeight17  =   21
      ButtonUseMaskColor17=   0   'False
      ButtonEnabled18 =   0   'False
      ButtonKey18     =   "18"
      BeginProperty ButtonFont18 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState18   =   5
      ButtonLeft18    =   843
      ButtonTop18     =   2
      ButtonWidth18   =   24
      ButtonHeight18  =   24
      ButtonUseMaskColor18=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13350
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmContas_pagar.frx":12C32
         Count           =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   60
      TabIndex        =   65
      Top             =   9210
      Width           =   15195
      Begin VB.TextBox txtTotalDevolver 
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
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Total a devolver."
         Top             =   390
         Width           =   1550
      End
      Begin VB.TextBox txtTotalAntecipado 
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
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Total antecipado."
         Top             =   390
         Width           =   1550
      End
      Begin VB.TextBox txtTotalPagar 
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
         MaxLength       =   20
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Total a pagar."
         Top             =   390
         Width           =   1550
      End
      Begin VB.TextBox TotalContas 
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
         Left            =   13440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Total geral."
         Top             =   390
         Width           =   1560
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   180
         TabIndex        =   70
         Top             =   330
         Width           =   8445
         _ExtentX        =   14896
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
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total devolver"
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
         Left            =   12045
         TabIndex        =   82
         Top             =   180
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total antecipado"
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
         Left            =   10380
         TabIndex        =   80
         Top             =   180
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total a pagar"
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
         Left            =   8985
         TabIndex        =   79
         Top             =   180
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total geral"
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
         Left            =   13770
         TabIndex        =   66
         Top             =   180
         Width           =   1515
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtIDContasReceber 
      Height          =   285
      Left            =   3030
      TabIndex        =   64
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSComctlLib.ListView lst_contas 
      Height          =   3650
      Left            =   60
      TabIndex        =   36
      Top             =   4960
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6429
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. vencto."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo docto."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   8312
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "IDempresa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar contas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   535
      Left            =   60
      TabIndex        =   68
      Top             =   4410
      Width           =   15195
      Begin VB.ComboBox cmbAno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmContas_pagar.frx":1D04A
         Left            =   14250
         List            =   "frmContas_pagar.frx":1D04C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         ToolTipText     =   "Ano."
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton OptAteomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
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
         Left            =   1020
         TabIndex        =   50
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton OptDomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
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
         Left            =   150
         TabIndex        =   49
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   345
         Left            =   2160
         TabIndex        =   51
         Top             =   240
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   609
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         TabMinWidth     =   1439
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   13
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jan"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fev"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mar"
               Key             =   "Mar"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Abril"
               Key             =   "Abr"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Maio"
               Key             =   "Maio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jun"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jul"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ago"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Set"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Out"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Nov"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dez"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Vencidas"
               Key             =   "Vencidas"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmContas_Pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Pagar As Boolean 'OK
Public StrSql_Contas_Pagar As String 'OK
Public StrSql_Contas_PagarTotal As String 'OK
Public StrSql_Contas_Pagar_AntecTotal As String 'OK
Public StrSql_Contas_Pagar_DevTotal As String 'OK
Dim TBLISTA_Contas_Pagar As ADODB.Recordset 'OK
Public Filtro_Contas_Pagar_Func As String 'OK
Public Filtro_Contas_Pagar_FuncRel As String 'OK
Public FormulaRel_Contas_Pagar As String 'OK

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=zKGGe6O5OUw&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=49&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_antecipacao_Click()
On Error GoTo tratar_erro

If Chk_antecipacao.Value = 1 Then Chk_devolucao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_devolucao_Click()
On Error GoTo tratar_erro

If Chk_devolucao.Value = 1 Then Chk_antecipacao.Value = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
ProcCarregaComboBanco

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
With lst_contas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) conta(s) antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmContas_pagar_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPlanoContas()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = True
Financeiro_Contas_Receber = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Recebidas = False
If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de cadastrar a(s) conta(s) contábil."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmFamilia_financeiro.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCC()
On Error GoTo tratar_erro

If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de cadastrar/visualizar o(s) centro(s) de custo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Financeiro_Contas_Pagar = True
Financeiro_Contas_Pagas = False
Faturamento = False
Permitido = True
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select ID_nota, txt_ndocumento, Txt_pedido from tbl_contaspagar where IdIntConta = " & IIf(txtidintconta = "", 0, txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    INNERJOINTEXTO = "NF.TipoNF from ((((tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID) INNER JOIN Compras_pedido CP ON CP.Pedido = PN.Proposta) INNER JOIN Compras_pedido_lista CPL ON CPL.IDpedido = CP.IDpedido) LEFT JOIN Compras_pedido_lista_custo CPLC ON CP.IDPedido = CPLC.IDpedido) LEFT JOIN Projproduto P ON P.Desenho = CPL.Desenho"
    If IsNull(TBContas!ID_nota) = False And TBContas!ID_nota <> "" And TBContas!ID_nota <> "0" Then
        TextoFiltro = "PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 2"
    ElseIf IsNull(TBContas!txt_ndocumento) = False And TBContas!txt_ndocumento <> "" Then
            TextoFiltro = "PN.NF = '" & TBContas!txt_ndocumento & "' and NF.int_TipoNota = 2"
        ElseIf IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" And TBContas!Txt_pedido <> "0" Then
                INNERJOINTEXTO = "CP.IDpedido from ((Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CPL.IDpedido = CP.IDpedido) LEFT JOIN Compras_pedido_lista_custo CPLC ON CP.IDPedido = CPLC.IDpedido) INNER JOIN Projproduto P ON P.Desenho = CPL.Desenho"
                TextoFiltro = "CP.Pedido = '" & TBContas!Txt_pedido & "'"
    End If
    Set TBProposta = CreateObject("adodb.recordset")
    TBProposta.Open "Select " & INNERJOINTEXTO & " where " & TextoFiltro & " and (CPLC.ID IS NOT NULL or P.Estoque = 'True')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProposta.EOF = False Then
        Permitido = False
    End If
End If
If Permitido = True Then frmContas_CC.Show 1 Else frmContas_pagar_lista_CC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With lst_contas
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(10) = 5
        .ButtonState(11) = 5
        .ButtonState(14) = 5
    ElseIf Cmb_opcao_lista = "Baixar" Then
            .ButtonState(4) = 5
            .ButtonState(10) = 0
            .ButtonState(11) = 5
            .ButtonState(14) = 5
        ElseIf Cmb_opcao_lista = "Status" Then
                .ButtonState(4) = 5
                .ButtonState(10) = 5
                .ButtonState(11) = 0
                .ButtonState(14) = 5
            ElseIf Cmb_opcao_lista = "Agendar" Then
                    .ButtonState(4) = 5
                    .ButtonState(10) = 5
                    .ButtonState(11) = 5
                    .ButtonState(14) = 0
                Else
                    .ButtonState(4) = 5
                    .ButtonState(10) = 5
                    .ButtonState(11) = 5
                    .ButtonState(14) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipo_Click()
On Error GoTo tratar_erro

txtIDFornec = ""
txtFornec = ""
If Cmb_tipo = "Cliente" Or Cmb_tipo = "Fornecedor" Then Cmd_localizar_contatos.Enabled = True Else Cmd_localizar_contatos.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_competencia_Click()
On Error GoTo tratar_erro

If txt_Competencia <> "" Then
    ProcFiltrarContas "Competencia = '" & txt_Competencia.Text & "'", "{tbl_ContasPagar.Competencia} = '" & txt_Competencia.Text & "'", True, False, False, False, Date, Date, "dt_Pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_data_transacao_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "CP.Data_transacao = '" & Format(Txt_data_transacao.Value, "Short Date") & "'", "{tbl_ContasPagar.Data_transacao} = Date(" & Year(Txt_data_transacao.Value) & "," & Month(Txt_data_transacao.Value) & "," & Day(Txt_data_transacao.Value) & ")", True, True, False, False, Txt_data_transacao, Txt_data_transacao, "Data_transacao"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcFiltrarContas(TextoFiltro As String, TextoFiltroRel As String, Imprimir As Boolean, DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataInicio As Date, DataFinal As Date, Ordenar As String)
On Error GoTo tratar_erro

NomeRel = "Contas_pagar.rpt"
ProcConstruirFiltroPadrao TextoFiltro, TextoFiltroRel, True, True
ProcSalvarDadosRel DataTransacao, DataEmissao, DataVencimento, DataInicio, DataFinal
StrSql_Contas_Pagar_AntecTotal = ""
StrSql_Contas_Pagar_DevTotal = ""
ProcCarregaLista (1)
Imprimir = Imprimir
frmContas_pagar_localizar.Todas_contas = False
'Novo_Pagar = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_contatos_Click()
On Error GoTo tratar_erro

If txtFornec <> "" Then
    Financeiro_Contas_Pagar = True
    Financeiro_Contas_Pagas = False
    Financeiro_Contas_Receber = False
    Financeiro_Contas_Recebidas = False
    If Cmb_tipo = "Fornecedor" Then
        Compras_Cotacao = False
        Compras_Pedido = False
        frmCompras_Pedido_contatos.Show 1
    Else
        Analise_critica = False
        Vendas_Proposta = False
        Vendas_PI = False
        Telemarketing = False
        Qualidade_PPAP_PSW = False
        frmVendas_propostaII_contato.Show 1
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_localizar_tipo_dcto_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = True
Financeiro_Contas_Receber = False
Clientes = False
Compras_Fornecedores = False
frmContas_Tipo_Dcto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_valor_Click()
On Error GoTo tratar_erro

If txtValorTotal <> "" Then
    valor = txtValorTotal
    NovoValor = Replace(valor, ",", ".")
    ProcFiltrarContas "dbl_valorpagto = " & NovoValor, "{tbl_ContasPagar.dbl_valorpagto} = " & NovoValor, True, False, False, False, Date, Date, "dt_Pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdForma_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = True
Financeiro_Forma_Pgto_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Forma_Pgto_Receber = False
frmContas_Forma_Pagamento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmdlocalizarforn_Click()
On Error GoTo tratar_erro

ProcLocalizarFornecedor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizarFornecedor()
On Error GoTo tratar_erro

ProcConfVariaveisLocCliente False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False
ProcConfVariaveisLocForn False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False
If Cmb_tipo = "Cliente" Then
    frmVendas_LocalizarCliente.Show 1
ElseIf Cmb_tipo = "Fornecedor" Then
        FrmCompras_localizafornecedor.Show 1
    ElseIf Cmb_tipo = "Funcionário" Then
            frmContas_pagar_localizar_func.Show 1
        Else
            frmContas_pagar_localizar_inst.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Pagar.AbsolutePage <> 2 Then
    If TBLISTA_Contas_Pagar.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Contas_Pagar.PageCount - 1)
    Else
        TBLISTA_Contas_Pagar.AbsolutePage = TBLISTA_Contas_Pagar.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Contas_Pagar.AbsolutePage)
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
    TBLISTA_Contas_Pagar.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Contas_Pagar.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Pagar.AbsolutePage = 1
ProcExibePagina (TBLISTA_Contas_Pagar.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Pagar.AbsolutePage <> -3 Then
    If TBLISTA_Contas_Pagar.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Contas_Pagar.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Contas_Pagar.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Pagar.AbsolutePage = TBLISTA_Contas_Pagar.PageCount
ProcExibePagina (TBLISTA_Contas_Pagar.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
On Error GoTo tratar_erro

If txtStatus.Text <> "" Then
    ProcFiltrarContas "Status = '" & txtStatus.Text & "'", "{tbl_ContasPagar.Status} = '" & txtStatus.Text & "'", True, False, False, False, Date, Date, "dt_Pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdtipo_Click()
On Error GoTo tratar_erro

If cmbtipo_conta.Text <> "" Then
    ProcFiltrarContas "CLASS_CONTA = '" & cmbtipo_conta.Text & "'", "{tbl_ContasPagar.CLASS_CONTA} = '" & cmbtipo_conta & "'", True, False, False, False, Date, Date, "dt_Pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF6: ProcPlanoContas
    Case vbKeyF7: ProcAgendaDia
    Case vbKeyF8: ProcParcelar
    Case vbKeyF9: ProcCopiar
    Case vbKeyF10: If Cmb_opcao_lista = "Baixar" Then ProcPagar
    Case vbKeyF11: If Cmb_opcao_lista = "Status" Then ProcStatus
    Case vbKeyF12: ProcCC
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmddoc_Click()
On Error GoTo tratar_erro

If txtNDocumento.Text <> "" Then
    ProcFiltrarContas "txt_NDocumento = '" & txtNDocumento.Text & "'", "{tbl_ContasPagar.txt_NDocumento} = '" & txtNDocumento.Text & "'", True, False, False, False, Date, Date, "dt_Pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarTodas()
On Error GoTo tratar_erro

NomeRel = "Contas_pagar.rpt"
ProcConstruirFiltroPadrao "CP.IDintconta IS NOT NULL", "Not(IsNull({tbl_ContasPagar.IDintconta}))", True, True
ProcSalvarDadosRel False, False, False, Date, Date
StrSql_Contas_Pagar_AntecTotal = ""
StrSql_Contas_Pagar_DevTotal = ""
ProcCarregaLista (1)
Imprimir = True
frmContas_pagar_localizar.Todas_contas = True
'Novo_Pagar = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdemissao_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "dt_emissao = '" & Format(txtDTEmissao.Value, "Short Date") & "'", "{tbl_ContasPagar.dt_emissao} = Date(" & Year(txtDTEmissao.Value) & "," & Month(txtDTEmissao.Value) & "," & Day(txtDTEmissao.Value) & ")", True, False, True, False, txtDTEmissao, txtDTEmissao, "dt_Emissao"

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
If txtidintconta = "" And Novo_Pagar = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If cmbtipo_conta = "" Then
    NomeCampo = "o tipo do documento"
    ProcVerificaAcao
    cmbtipo_conta.SetFocus
    Exit Sub
End If
If txtNDocumento.Text = "" Then
    NomeCampo = "o número do documento"
    ProcVerificaAcao
    txtNDocumento.SetFocus
    Exit Sub
End If
txtparcela.PromptInclude = False
If Len(txtparcela) < 6 Then
    USMsgBox ("O número da parcela digitada não é válido, digite o número correto."), vbExclamation, "CAPRIND v5.0"
    txtparcela.SetFocus
    Exit Sub
End If
txtparcela.PromptInclude = True

If txtIDFornec.Text = "" Or txtIDFornec.Text = "" Then
    NomeCampo = "o fornecedor"
    ProcVerificaAcao
    Cmdlocalizarforn_Click
    Exit Sub
End If

valor = IIf(txtValorTotal = "", 0, txtValorTotal)
If Chk_devolucao.Value = 1 And valor >= 0 Then
    txtValorTotal = Format(txtValorTotal, "-###,##0.00")
ElseIf Chk_devolucao.Value = 0 And valor <= 0 Then
        NomeCampo = "o valor"
        ProcVerificaAcao
        txtValorTotal.SetFocus
        Exit Sub
End If

'Verifica se é antecipação e se já foi vinculado em alguma conta paga
If Chk_antecipacao.Value = 1 Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_Contas_antecipacao where ID_antecipacao = " & IIf(txtidintconta = "", 0, txtidintconta) & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        USMsgBox ("Não é permitido salvar, pois esta antecipação já esta relacionada a uma conta baixada."), vbExclamation, "CAPRIND v5.0"
        TBContas.Close
        Exit Sub
    End If
End If

If Cmb_tipo = "Cliente" Then
    Tipo = "CL"
    IDforn = txtIDFornec
ElseIf Cmb_tipo = "Fornecedor" Then
        Tipo = "FO"
        IDforn = txtIDFornec
    ElseIf Cmb_tipo = "Funcionário" Then
            Tipo = "FU"
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select ID from Funcionarios where Codigo = '" & txtIDFornec & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                IDforn = TBAbrir!ID
            End If
            TBAbrir.Close
        Else
            Tipo = "IN"
            IDforn = txtIDFornec
End If

'Verifica se já existe conta com o mesmo número de doc./nf e vencimento para o fornecedor
If Novo_Pagar = True Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_contaspagar where txt_NDocumento = '" & txtNDocumento & "' and dt_Pagamento = '" & txtDtpagto & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        If TBContas!int_codforn <> IDforn Then
            USMsgBox ("Já existe uma conta cadastrada com este número de documento " & txtNDocumento & " com vencimento em " & txtDtpagto & " para o fornecedor " & TBContas!Txt_fornecedor & "."), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Já existe uma conta cadastrada com este número de documento " & txtNDocumento & " com vencimento em " & txtDtpagto & "  para este fornecedor."), vbExclamation, "CAPRIND v5.0"
        End If
        If USMsgBox("Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
            TBContas.Close
            Exit Sub
        End If
    End If
End If

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contaspagar where IdIntConta = " & IIf(txtidintconta = "", 0, txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    If TBContas!status <> "TÍTULO EM ABERTO" Then
        USMsgBox ("Não é permitido alterar esta conta, pois a mesma já foi baixada parcial, está bloqueada ou é uma antecipação baixada."), vbExclamation, "CAPRIND v5.0"
        TBContas.Close
        Exit Sub
    End If
    'Corrige o valor das contas contábeis
    If TBContas!dbl_valorpagto <> valor And Lista_PC.ListItems.Count <> 0 Then
        If USMsgBox("Deseja atualizar o valor da(s) conta(s) contábil(eis)?)", vbYesNo) = vbYes Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Valor from Familia_financeiro where IDConta = " & IIf(txtidintconta = "", 0, txtidintconta) & " and TipoConta = 'P' and Valor > 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Valor1 = (TBAbrir!valor / TBContas!dbl_valorpagto) * 100
                    TBAbrir!valor = (valor * Valor1) / 100
                    TBAbrir.Update
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
            ProcCarregaListaPC
        End If
    End If
    If TBContas!dbl_valorpagto <> valor Or TBContas!dt_Pagamento <> txtDtpagto Then
    
        'Corrige o valor e a data dos centros de custo
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CC_realizado where ID_financeiro = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                TBFI!Data = txtDtpagto
                TBFI!valor = Format((valor * TBFI!Percentual) / 100, "###,##0.00")
                TBFI.Update
        
                'Grava movimentação no centro consolidado
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFI!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    Do While TBAfericao.EOF = False
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select * from CC_realizado where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = False Then
                            TBGravar!Data = txtDtpagto
                            TBGravar!valor = Format((valor * TBGravar!Percentual) / 100, "###,##0.00")
                            TBGravar.Update
                        End If
                        TBGravar.Close
                        
                        Set TBCiclo = CreateObject("adodb.recordset")
                        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCiclo.EOF = False Then
                            Do While TBCiclo.EOF = False
                                Set TBGravar = CreateObject("adodb.recordset")
                                TBGravar.Open "Select * from CC_realizado where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                                If TBGravar.EOF = False Then
                                    TBGravar!Data = txtDtpagto
                                    TBGravar!valor = Format((valor * TBGravar!Percentual) / 100, "###,##0.00")
                                    TBGravar.Update
                                End If
                                TBGravar.Close
                                
                                TBCiclo.MoveNext
                            Loop
                        End If
                        TBCiclo.Close
                        
                        TBAfericao.MoveNext
                    Loop
                End If
                TBAfericao.Close
                
                TBFI.MoveNext
            Loop
        End If
        TBFI.Close
    End If
Else
    TBContas.AddNew
    TBContas!Despesas_NF = False
    TBContas!Parcial = False
    TBContas!impresso = False
    TBContas!Bloqueado = False
    TBContas!Logsit = "N"
    TBContas!Responsavel = pubUsuario
End If
If Chk_antecipacao.Value = 1 Then
    TBContas!Antecipacao = True
    TBContas!Saldo_antecipacao = txtValorTotal.Text
Else
    TBContas!Antecipacao = False
End If
If Chk_devolucao.Value = 1 Then TBContas!Devolucao = True Else TBContas!Devolucao = False
If Chk_agendado.Value = 1 Then TBContas!Agendado = True Else TBContas!Agendado = False
If chkConta_fixa.Value = 1 Then TBContas!Conta_fixa = True Else TBContas!Conta_fixa = False

TBContas!Data_transacao = Txt_data_transacao.Value
TBContas!Dt_emissao = txtDTEmissao.Value
TBContas!dt_Pagamento = txtDtpagto.Value
TBContas!dbl_valorpagto = txtValorTotal.Text
TBContas!Banco = IIf(cmbBanco = "", Null, cmbBanco)
TBContas!FormaBaixa = cmb_forma
TBContas!Competencia = txt_Competencia
TBContas!txt_observacoes = txtObs.Text
TBContas!Txt_pedido = txtNPedido.Text

TBContas!Tipo = Tipo
TBContas!int_codforn = IDforn
TBContas!Txt_fornecedor = txtFornec.Text

TBContas!txt_ndocumento = txtNDocumento.Text
TBContas!Class_conta = cmbtipo_conta.Text
TBContas!txt_Parcela = txtparcela.Text
TBContas!status = txtStatus
TBContas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

TBContas.Update
txtidintconta = TBContas!IDintconta

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = txtidintconta
TBFluxo!Operacao = "À Debitar"
TBFluxo!Data = txtDtpagto
TBFluxo!valor = txtValorTotal
TBFluxo!Descricao = txtFornec
TBFluxo!status = "N"
TBFluxo!int_NotaFiscal = txtNDocumento
TBFluxo!Instituicao = IIf(cmbBanco = "", Null, cmbBanco)
TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "UPDATE tbl_contaspagar set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & IIf(txtidintconta = "", 0, txtidintconta)
TBFluxo.Close

TBContas.Close
If Novo_Pagar = True Then
    USMsgBox ("Nova conta cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
    ProcConstruirFiltroPadrao "CP.IdIntConta = " & txtidintconta, "{tbl_ContasPagar.IDintconta} = " & txtidintconta, True, True
    ProcCarregaLista (1)
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    ProcCarregaLista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And lst_contas.ListItems.Count <> 0 Then
        lst_contas.SelectedItem = lst_contas.ListItems(CodigoLista)
        lst_contas.SetFocus
    End If
End If

1:
    '==================================
    Modulo = "Financeiro/Contas a pagar"
    ID_documento = txtidintconta
    Documento = "Documento: " & txtNDocumento
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_Pagar = False
    
    'Verifica se a empresa exige centro de custo
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select IdIntConta, ID_nota, Txt_pedido from tbl_contaspagar where IdIntConta = " & IIf(txtidintconta = "", 0, txtidintconta), Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        If (IsNull(TBContas!ID_nota) = True Or TBContas!ID_nota = "" Or TBContas!ID_nota = "0") And (IsNull(TBContas!Txt_pedido) = True Or TBContas!Txt_pedido = "" Or TBContas!Txt_pedido = "0") Then
            If ProcVerifExigeCC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "ID_financeiro = " & txtidintconta, True) = False Then ProcCC
        End If
    End If
    TBContas.Close

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizar_fornecedor_Click()
On Error GoTo tratar_erro
    
If txtFornec.Text <> "" Then
    ProcFiltrarContas "txt_fornecedor = '" & txtFornec.Text & "'", "{tbl_ContasPagar.txt_fornecedor} = '" & txtFornec.Text & "'", True, False, False, False, Date, Date, "dt_Pagamento"
Else
    ProcFiltrarTodas
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

Proposta = True
If txtNPedido.Text <> "" Then
    NomeRel = "Contas_pagar.rpt"
    ProcConstruirFiltroPadrao "PN.Proposta = '" & txtNPedido & "'", "{tbl_proposta_nota.proposta} = '" & txtNPedido & "'", True, True
    ProcSalvarDadosRel False, False, False, Date, Date
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open StrSql_Contas_Pagar, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        ProcConstruirFiltroPadrao "CP.txt_pedido = '" & txtNPedido & "'", "{tbl_ContasPagar.txt_pedido} = '" & txtNPedido & "'", True, True
    End If
    TBAbrir.Close
    Imprimir = True
    frmContas_pagar_localizar.Todas_contas = False
Else
    ProcFiltrarTodas
End If
StrSql_Contas_Pagar_AntecTotal = ""
StrSql_Contas_Pagar_DevTotal = ""
ProcCarregaLista (1)
Proposta = False
Novo_Pagar = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcConstruirFiltroPadrao(TextoFiltro As String, TextoFiltroRel As String, ApagarFiltroAntec As Boolean, ApagarFiltroDev As Boolean)
On Error GoTo tratar_erro

CamposFiltro = "CP.IDintconta, CP.Dt_emissao, CP.dt_Pagamento, CP.dbl_valorpagto, CP.Class_conta, CP.txt_ndocumento, CP.txt_Parcela, CP.Txt_fornecedor, CP.ID_empresa, CP.Responsavel, CP.Antecipacao, CP.Saldo_antecipacao"
If Left(TextoFiltro, 2) = "PN" Then INNERJOINPADRAO = " from tbl_ContasPagar CP INNER JOIN tbl_proposta_nota PN ON PN.ID_nota = CP.ID_nota" Else INNERJOINPADRAO = " from tbl_ContasPagar CP"
INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
INNERJOINTEXTOSUM = "Select SUM(CP.dbl_valorpagto) AS TotContas " & INNERJOINPADRAO
TextoFiltroPadrao = "CP.logsit = 'N' and CP.bloqueado = 'False' and CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & IIf(ApagarFiltroAntec = True, " and CP.Antecipacao = 'False'", "") & IIf(ApagarFiltroDev = True, " and CP.Devolucao = 'False'", "") & " and " & Filtro_Contas_Pagar_Func
TextoFiltroPadraoRel = "{tbl_ContasPagar.LogSit} = 'N' and {tbl_ContasPagar.bloqueado} = False and {tbl_ContasPagar.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & IIf(ApagarFiltroAntec = True, " and {tbl_ContasPagar.Antecipacao} = False", "") & IIf(ApagarFiltroDev = True, " and {tbl_ContasPagar.Devolucao} = False", "") & " and " & Filtro_Contas_Pagar_FuncRel
OrdenarTexto = " group by " & CamposFiltro & " order by CP.dt_Pagamento, CP.IdIntConta"
StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadrao
If ApagarFiltroAntec = True Then StrSql_Contas_Pagar_AntecTotal = ""
If ApagarFiltroDev = True Then StrSql_Contas_Pagar_DevTotal = ""
FormulaRel_Contas_Pagar = TextoFiltroRel & " and " & TextoFiltroPadraoRel

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdvencimento_Click()
On Error GoTo tratar_erro

ProcFiltrarContas "dt_pagamento = '" & Format(txtDtpagto.Value, "Short Date") & "'", "{tbl_ContasPagar.dt_pagamento} = Date(" & Year(txtDtpagto.Value) & "," & Month(txtDtpagto.Value) & "," & Day(txtDtpagto.Value) & ")", True, False, False, True, txtDtpagto, txtDtpagto, "dt_Pagamento"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 17, True

Formulario = "Financeiro/Contas a pagar"
Direitos
ProcLimpaVariaveisPrincipais
Cmb_opcao_lista = "Baixar"
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
txtDTEmissao.Value = Date
txtDtpagto.Value = Date
ProcVerifAcessosContasFunc
ProcCarregaComboAno cmbAno, Year(Now) - 2, 2
TabFiltro.Tabs(Month(Date)).Selected = True

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifAcessosContasFunc()
On Error GoTo tratar_erro

'Verifica se o usuário pode visualizar as contas dos funcionários
If FunVerifAcessoContasFunc("Financeiro/Contas a pagar/Visualizar contas dos funcionários") = True Then
    Filtro_Contas_Pagar_Func = "CP.txt_fornecedor <> 'Null'"
    Filtro_Contas_Pagar_FuncRel = "{tbl_contaspagar.txt_fornecedor} <> 'Null'"
Else
    Filtro_Contas_Pagar_Func = "CP.Tipo <> 'FU'"
    Filtro_Contas_Pagar_FuncRel = "{tbl_contaspagar.Tipo} <> 'FU'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboBanco()
On Error GoTo tratar_erro

ProcCarregaComboBancoFinanceiro cmbBanco, "txt_Descricao IS NOT NULL and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Bloqueado = 'false' and DtValidacao IS NOT NULL", True
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

ProcCarregaComboFormaPgtoRcbto cmb_forma, "Tipo = 'P'"
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaTipoDocumento()
On Error GoTo tratar_erro

ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'P'"
If txtidintconta <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmbBanco.ListIndex = -1
cmb_forma.ListIndex = -1
cmbtipo_conta.ListIndex = -1
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_ContasPagar where IdIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If IsNull(TBAbrir!Banco) = False And TBAbrir!Banco <> "" Then cmbBanco = TBAbrir!Banco
    If IsNull(TBAbrir!FormaBaixa) = False And TBAbrir!FormaBaixa <> "" Then cmb_forma = TBAbrir!FormaBaixa
    If IsNull(TBAbrir!Class_conta) = False And TBAbrir!Class_conta <> "" Then cmbtipo_conta = TBAbrir!Class_conta
1:
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtidintconta.Text = ""
Txt_data_transacao.Value = Date
txtNDocumento.Text = ""
txtNPedido.Clear
txtStatus.Text = "TÍTULO EM ABERTO"
txtparcela.Text = "___/___"
txtDTEmissao.Value = Date
cmbtipo_conta.ListIndex = -1
txtDtpagto.Value = Date
txtValorTotal.Text = ""
txtIDFornec.Text = ""
Cmb_tipo = "Fornecedor"
txtFornec.Text = ""
cmbBanco.ListIndex = -1
cmb_forma.ListIndex = -1
txt_Competencia = ""
Chk_antecipacao.Value = 0
Chk_devolucao.Value = 0
Chk_agendado.Value = 0
chkConta_fixa.Value = 0
txtObs.Text = ""
CodigoLista = 0
Lista_PC.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista(Pagina As Integer)
On Error GoTo tratar_erro

lst_contas.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Set TBLISTA_Contas_Pagar = CreateObject("adodb.recordset")
TBLISTA_Contas_Pagar.Open StrSql_Contas_Pagar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Contas_Pagar.EOF = False Then ProcExibePagina (Pagina)
ProcCarregaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Codproduto = 0
Dataini = 0
TotContas = 0
lst_contas.ListItems.Clear
TBLISTA_Contas_Pagar.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Contas_Pagar.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Contas_Pagar.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Contas_Pagar.RecordCount - IIf(Pagina > 1, (TBLISTA_Contas_Pagar.PageSize * (Pagina - 1)), 0), TBLISTA_Contas_Pagar.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Contas_Pagar.EOF = False And (ContadorReg <= TamanhoPagina)
    With lst_contas.ListItems.Add(, , TBLISTA_Contas_Pagar!IDintconta)
        .SubItems(1) = Format(TBLISTA_Contas_Pagar!Dt_emissao, "dd/mm/yy")
        .SubItems(2) = Format(TBLISTA_Contas_Pagar!dt_Pagamento, "dd/mm/yy")
        
        If TBLISTA_Contas_Pagar!Antecipacao = True Then qt = IIf(IsNull(TBLISTA_Contas_Pagar!Saldo_antecipacao), 0, TBLISTA_Contas_Pagar!Saldo_antecipacao) Else qt = IIf(IsNull(TBLISTA_Contas_Pagar!dbl_valorpagto), 0, TBLISTA_Contas_Pagar!dbl_valorpagto)
        .SubItems(3) = Format(qt, "###,##0.00")
        
        .SubItems(4) = IIf(IsNull(TBLISTA_Contas_Pagar!Class_conta), "", TBLISTA_Contas_Pagar!Class_conta)
        .SubItems(5) = IIf(IsNull(TBLISTA_Contas_Pagar!txt_ndocumento), "", TBLISTA_Contas_Pagar!txt_ndocumento)
        .SubItems(6) = IIf(IsNull(TBLISTA_Contas_Pagar!txt_Parcela), "", TBLISTA_Contas_Pagar!txt_Parcela)
        .SubItems(7) = IIf(IsNull(TBLISTA_Contas_Pagar!Txt_fornecedor), "", Trim(TBLISTA_Contas_Pagar!Txt_fornecedor))
        .SubItems(8) = IIf(IsNull(TBLISTA_Contas_Pagar!ID_empresa), 0, TBLISTA_Contas_Pagar!ID_empresa)
        .SubItems(9) = IIf(IsNull(TBLISTA_Contas_Pagar!Responsavel), 0, TBLISTA_Contas_Pagar!Responsavel)
        Dataini = TBLISTA_Contas_Pagar!dt_Pagamento
        If Date > Dataini Then
            .ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .ListSubItems(3).ForeColor = vbRed
            .ListSubItems(4).ForeColor = vbRed
            .ListSubItems(5).ForeColor = vbRed
            .ListSubItems(6).ForeColor = vbRed
            .ListSubItems(7).ForeColor = vbRed
            .ListSubItems(8).ForeColor = vbRed
            .ListSubItems(9).ForeColor = vbRed
        End If
    End With
    TBLISTA_Contas_Pagar.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Contas_Pagar.RecordCount
If TBLISTA_Contas_Pagar.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Contas_Pagar.PageCount
ElseIf TBLISTA_Contas_Pagar.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Pagar.PageCount & " de: " & TBLISTA_Contas_Pagar.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Contas_Pagar.AbsolutePage - 1 & " de: " & TBLISTA_Contas_Pagar.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaTotal()
On Error GoTo tratar_erro

'À pagar
valor = 0
Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open StrSql_Contas_PagarTotal, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    valor = IIf(IsNull(TBTotaisnota!TotContas), 0, TBTotaisnota!TotContas)
End If
TBTotaisnota.Close
txtTotalPagar.Text = Format(valor, "###,##0.00")

'Antecipado
Valor1 = 0
If StrSql_Contas_Pagar_AntecTotal <> "" Then
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open StrSql_Contas_Pagar_AntecTotal, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = False Then
        Valor1 = IIf(IsNull(TBTotaisnota!TotContas1), 0, TBTotaisnota!TotContas1)
    End If
End If
txtTotalAntecipado.Text = Format(Valor1, "###,##0.00")

'Devolver
Valor2 = 0
If StrSql_Contas_Pagar_DevTotal <> "" Then
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open StrSql_Contas_Pagar_DevTotal, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = False Then
        Valor2 = IIf(IsNull(TBTotaisnota!TotContas), 0, TBTotaisnota!TotContas)
    End If
End If
txtTotalDevolver.Text = Format(Valor2, "###,##0.00")

'Total geral (A pagar - Antecipado + A devolver)
qt = valor - Valor1 + Valor2
TotalContas.Text = IIf(qt < 0, "0,00", Format(qt, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Contas a pagar"
Direitos
ProcCarregaTipoDocumento
ProcCarregaComboBanco
ProcCarregaComboForma
ProcLimpaVariaveisPrincipais
ProcVerifAcessosContasFunc
NomeRel = "Contas_pagar.rpt"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frmContas_pagar_localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcParcelar()
On Error GoTo tratar_erro
    
If txtNDocumento.Text = "" Or txtValorTotal.Text = "" Or txtIDFornec.Text = "" Then
    USMsgBox ("Informe a conta antes de parcelar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Pagar = True Then
    USMsgBox ("Salve a conta antes de parcelar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Chk_antecipacao.Value = 1 Then
    USMsgBox ("Não é permitido parcelar esta conta, pois a mesma é uma antecipação."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
GerarPagtos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAgendaDia()
On Error GoTo tratar_erro

ProcFiltrarContas "dt_pagamento = '" & Format(Date, "Short Date") & "'", "{tbl_ContasPagar.dt_pagamento} = Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")", True, False, False, False, Date, Date, "dt_Pagamento"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPagar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Contador = 0
With lst_contas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                'Verifica a primeira conta selecionada e informa quais as novas contas poderão ser selecionadas
                If Contador = 0 Then
                    If TBContas!Antecipacao = True Then
                        Contador = 1
                    ElseIf TBContas!Devolucao = True Then
                            Contador = 2
                        Else
                            Contador = 3
                    End If
                ElseIf Contador = 1 Then
                        If TBContas!Antecipacao = False Then
                            USMsgBox ("Só é permitido baixar conta de antecipação."), vbExclamation, "CAPRIND v5.0"
                            Exit Sub
                        End If
                    ElseIf Contador = 2 Then
                            If TBContas!Devolucao = False Then
                                USMsgBox ("Só é permitido baixar conta de devolução."), vbExclamation, "CAPRIND v5.0"
                                Exit Sub
                            End If
                        Else
                            If TBContas!status = "DUPLICATA DESCONTADA EM ABERTO" Or TBContas!Antecipacao = True Or TBContas!Devolucao = True Then
                                USMsgBox ("Só é permitido baixar conta em aberto ou baixada parcial que não seja conta de antecipação nem devolução."), vbExclamation, "CAPRIND v5.0"
                                Exit Sub
                            End If
                End If
            End If
            TBContas.Close
        End If
    Next InitFor
End With

Permitido1 = False
With lst_contas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido1 = False Then
                If USMsgBox("Deseja realmente baixar esta(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido1 = True
            GoTo 2
        End If
    Next InitFor
End With
2:
    If Permitido1 = False Then
        USMsgBox ("Informe a(s) conta(s) antes de baixar."), vbExclamation, "CAPRIND v5.0"
    Else
        Permitido1 = False
        frm_Baixas.Show 1
        If Permitido1 = True Then
            ProcCarregaLista (1)
            ProcLimpaCampos
            lst_contas.SetFocus
            If lst_contas.ListItems.Count = 0 Then Exit Sub
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & lst_contas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                ProcCarregaDados
                CodigoLista = lst_contas.SelectedItem.index
            End If
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
With lst_contas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                'Fluxo de Caixa
                Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo)
                
                If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                    'Fluxo de Caixa
                    If TBContas!FormaBaixa = "CHEQUE" Or TBContas!FormaBaixa = "CHEQUE PRÉ-DATADO" Or TBContas!FormaBaixa = "DOC" Or TBContas!FormaBaixa = "TED" Or TBContas!FormaBaixa = "MALOTE" Or IsNull(TBContas!ID_varias) = False And TBContas!ID_varias > 0 Then
                        TextoFiltroData = "Data = '" & Format(TBContas!Data_movimentacao, "Short Date") & "' and"
                        Select Case TBContas!FormaBaixa
                            Case "CHEQUE":
                                Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                                TextoFiltroData = ""
                            Case "CHEQUE PRÉ-DATADO":
                                Cheque = "Cheque n. " & TBContas!NDoctoBaixa
                                TextoFiltroData = ""
                            Case "DOC": Cheque = "Doc n. " & TBContas!NDoctoBaixa
                            Case "TED": Cheque = "Ted n. " & TBContas!NDoctoBaixa
                            Case "MALOTE": Cheque = "Malote n. " & TBContas!NDoctoBaixa
                        End Select
                        Set TBFluxo = CreateObject("adodb.recordset")
                        If Left(TBContas!FormaBaixa, 6) = "CHEQUE" Or TBContas!FormaBaixa = "DOC" Or TBContas!FormaBaixa = "TED" Or TBContas!FormaBaixa = "MALOTE" Then
                            TextoFiltro = TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Débito' and (idintconta = 0 or idintconta IS NULL)"
                        Else
                            If IsNull(TBContas!ID_varias) = True Or TBContas!ID_varias = 0 Then TextoFiltro = TextoFiltroData & " Instituicao = '" & TBContas!Banco & "' and Descricao = '" & Cheque & "' and Operacao = 'Débito'" Else TextoFiltro = "ID_varias = " & TBContas!ID_varias
                        End If
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = False Then
                            TBFluxo!valor = Format(TBFluxo!valor - TBContas!dbl_valorpagto, "###,##0.00")
                            TBFluxo.Update
                            If TBFluxo!valor <= 0 Then
                                TBFluxo.Delete
                                Conexao.Execute "DELETE from tbl_Contas_Varias where ID = " & IIf(IsNull(TBContas!ID_varias), 0, TBContas!ID_varias)
                            End If
                        End If
                    End If
                    
                    If TBContas!FormaBaixa = "SAQUE" Then
                        'Verifica saque e atualiza saldo
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select IDSaque from tbl_ContasPagar_Saque where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBProduto.EOF = False Then
                            ProcAtualizaSaldoSaque TBProduto!IDSaque
                        End If
                        TBProduto.Close
                        Conexao.Execute "DELETE from tbl_ContasPagar_Saque where IdIntConta = " & .ListItems(InitFor)
                    ElseIf TBContas!FormaBaixa <> "CHEQUE" And TBContas!FormaBaixa <> "CHEQUE PRÉ-DATADO" Then
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select * from tbl_instituicoes where txt_descricao = '" & TBContas!Banco & "' and ID_empresa = " & TBContas!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then
                                TBProduto!Saldo = TBProduto!Saldo + TBContas!dbl_valorpagto
                                TBProduto.Update
                            End If
                            TBProduto.Close
                    End If
                End If
            End If
            TBContas.Close
            Conexao.Execute "DELETE FROM tbl_ContasPagar WHERE IdIntConta = " & .ListItems(InitFor)
            Conexao.Execute "DELETE FROM familia_financeiro WHERE IdConta = " & .ListItems(InitFor) & " and TipoConta = 'P' and Deposito_transf = 'False'"
            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_financeiro = " & .ListItems(InitFor) & " and ID_duplicata = 0"
            Conexao.Execute "UPDATE CC_realizado set ID_Financeiro = 0 where ID_financeiro = " & .ListItems(InitFor)
            
            '====================================
            Modulo = "Financeiro/Contas a pagar"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Documento: " & .ListItems(InitFor).SubItems(5)
            Documento1 = ""
            ProcGravaEvento
            '===================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) conta(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Conta(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    lst_contas.SetFocus
    Novo_Pagar = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAgendarPgto()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
With lst_contas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente alterar o agendamento de pagamento desta(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
1:
            Permitido = True
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select Banco, Agendado from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Agendado = True Then
                    TBContas!Agendado = False
                Else
                    With frmContas_pagar_agendar
                        If Permitido1 = False Then .Show 1
                        Permitido1 = True
                        TBContas!Banco = .IB
                        TBContas!Agendado = True
                    End With
                End If
                TBContas.Update
            End If
            TBContas.Close
            
            '====================================
            Modulo = "Financeiro/Contas a pagar"
            Evento = "Alterar o agendamento de pagamento"
            ID_documento = .ListItems(InitFor)
            Documento = "Documento: " & .ListItems(InitFor).SubItems(5)
            Documento1 = ""
            ProcGravaEvento
            '===================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) conta(s) antes de alterar o agendamento de pagamento."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Alteração(ões) efetuada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista (1)
    lst_contas.SetFocus
    Novo_Pagar = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If frmContas_pagar_localizar.Todas_contas = True Then
    frmContas_pagar_menuimpressao.Show 1
Else
    ProcVerificaContasSelRel lst_contas, IIf(Cmb_opcao_lista = "Relatório", True, False)
    If Familiatext <> "" Then
        ProcImprimirRel "(" & Familiatext & ") and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    Else
        ProcImprimirRel FormulaRel_Contas_Pagar & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    End If
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
Novo_Pagar = True
Chk_antecipacao.Enabled = True
Chk_devolucao.Enabled = True
Chk_agendado.Enabled = True
chkConta_fixa.Enabled = True
Txt_data_transacao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro
    
If txtNDocumento.Text = "" Or txtValorTotal.Text = "" Or txtIDFornec.Text = "" Then
    USMsgBox ("Informe a conta antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Pagar = True Then
    USMsgBox ("Salve a conta antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Chk_antecipacao.Value = 1 Then
    USMsgBox ("Não é permitido copiar esta conta, pois a mesma é uma antecipação."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frm_Contas_parcelamento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Pagar = True Then
    If USMsgBox("A conta ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Pagar = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Pagar = False
TotContas = 0
StrSql_Contas_Pagar = ""
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_PC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_PC, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_PC_DblClick()
On Error GoTo tratar_erro

If Lista_PC.ListItems.Count = 0 Then Exit Sub
Qtde = 0
Valor_conta = ""

Mensagem:
    Valor_conta = InputBox("Favor informar o novo valor da conta contábil.")
    If Valor_conta = "" Then Exit Sub
    If IsNumeric(Valor_conta) = False Then
        USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
        GoTo Mensagem
    End If
    Qtde = Valor_conta
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_contaspagar where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If TBAbrir!Devolucao = True And Qtde >= 0 Then
            Qtde = Format(Qtde, "-###,##0.00")
        ElseIf TBAbrir!Devolucao = False And Qtde <= 0 Then
                USMsgBox ("Informe o novo valor antes de alterar."), vbExclamation, "CAPRIND v5.0"
                GoTo Mensagem
        End If
    End If
    
    'Verifica saldo das contas contábeis
    valor = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(Valor) as Valor from Familia_financeiro where IDConta = " & txtidintconta & " and TipoConta = 'P' and ID <> " & Lista_PC.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select dbl_valorpagto, Devolucao from tbl_contaspagar where IDIntConta = " & txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qt = TBAbrir!dbl_valorpagto
        Permitido = True
        If TBAbrir!Devolucao = True Then
            If (valor + Qtde) < qt Then Permitido = False
        Else
            If (valor + Qtde) > qt Then Permitido = False
        End If
        If Permitido = False Then
            USMsgBox ("Não é permitido alterar, pois o valor digitado ultrapassa o saldo da conta."), vbExclamation, "CAPRIND v5.0"
            GoTo Mensagem
        End If
    End If
    TBAbrir.Close
    
    NovoValor = Replace(Qtde, ",", ".")
    Conexao.Execute "Update Familia_financeiro Set Valor = " & NovoValor & " where ID = " & Lista_PC.SelectedItem
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaPC
    '====================================
    Modulo = "Financeiro/Contas a pagar"
    Evento = "Alterar valor da conta contábil"
    ID_documento = Lista_PC.SelectedItem
    Documento = "Documento: " & txtNDocumento
    Documento1 = "Código do plano: " & Lista_PC.SelectedItem.ListSubItems(1) & " - Descrição do plano: " & Lista_PC.SelectedItem.ListSubItems(2)
    ProcGravaEvento
    '===================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lst_Contas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    Contador = 0
    With lst_contas
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If Cmb_opcao_lista = "Excluir" Then
                        If IsNull(TBContas!IdContaReceber) = False And TBContas!IdContaReceber <> "" And TBContas!IdContaReceber <> "0" Then GoTo Proximo
                        If TBContas!status <> "TÍTULO EM ABERTO" And TBContas!status <> "BLOQUEADA" And TBContas!status <> "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                        If TBContas!Despesas_NF = True Then GoTo Proximo
                    ElseIf Cmb_opcao_lista = "Status" Then
                            If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                        ElseIf Cmb_opcao_lista = "Baixar" Then
                                If TBContas!status = "BLOQUEADA" Or TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then GoTo Proximo
                                'Verifica se a empresa exige centro de custo
                                If (IsNull(TBContas!ID_nota) = True Or TBContas!ID_nota = "" Or TBContas!ID_nota = "0") And (IsNull(TBContas!Txt_pedido) = True Or TBContas!Txt_pedido = "" Or TBContas!Txt_pedido = "0") Then
                                    If ProcVerifExigeCC(TBContas!ID_empresa, "ID_financeiro = " & TBContas!IDintconta, False) = False Then GoTo Proximo
                                End If
                    End If
                End If
                TBContas.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lst_contas, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_contas_DblClick()
On Error GoTo tratar_erro

If lst_contas.ListItems.Count = 0 Then Exit Sub

TextoPedido = ""
TextoPedidoRel = ""
Contador2 = 0
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select ID_nota, Txt_pedido from tbl_ContasPagar where IdIntConta = " & lst_contas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Set TBProposta = CreateObject("adodb.recordset")
    TBProposta.Open "Select CP.IDpedido from (tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID) INNER JOIN Compras_pedido CP ON CP.Pedido = PN.Proposta where PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
    If TBProposta.EOF = False Then
        Do While TBProposta.EOF = False
            If TextoPedido = "" Then
                TextoPedido = "CP.IDpedido = " & TBProposta!IDpedido
                TextoPedidoRel = "{compras_pedido.IDpedido} = " & TBProposta!IDpedido
            Else
                TextoPedido = TextoPedido & " or CP.IDpedido = " & TBProposta!IDpedido
                TextoPedidoRel = TextoPedidoRel & " or {compras_pedido.IDpedido} = " & TBProposta!IDpedido
            End If
            Contador2 = Contador2 + 1
            TBProposta.MoveNext
        Loop
    ElseIf TBContas!Txt_pedido <> "" And IsNull(TBContas!Txt_pedido) = False Then
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select IDpedido from Compras_pedido where pedido = '" & TBContas!Txt_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Contador2 = 1
            TextoPedido = "CP.IDpedido = " & TBProposta!IDpedido
            TextoPedidoRel = "{compras_pedido.IDpedido} = " & TBProposta!IDpedido
        End If
    End If
    TBProposta.Close
End If
TBContas.Close

If TextoPedido <> "" Then
    Formulario = "Compras/Pedido"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    
    With frmCompras_Pedido
        .Sql_Pedido_Localizar = "Select CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado from Compras_pedido CP LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao where " & TextoPedido
        .FormulaRel_Pedido = TextoPedidoRel
        .listapedido.ListItems.Clear
        .ProcAtualizalistapedido (1)
        
        If Contador2 = 1 Then
            Set TBCompras_Pedido = CreateObject("adodb.recordset")
            TBCompras_Pedido.Open "Select * from compras_pedido CP where " & TextoPedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Pedido.EOF = False Then
                .ProcLimpar
                .ProcLimpaCamposItem True
                .ProcLimpaCamposServ True
                .ProcPuxaDados
            End If
            TBCompras_Pedido.Close
        End If
        .SSTab1.Tab = 0
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_contas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lst_contas
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & Item, Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If Cmb_opcao_lista = "Excluir" Then
                    If IsNull(TBContas!IdContaReceber) = False And TBContas!IdContaReceber <> "" And TBContas!IdContaReceber <> "0" Then
                        USMsgBox ("Não é permitido excluir esta conta, pois a mesma está vinculada a uma conta a receber descontada."), vbExclamation, "CAPRIND v5.0"
                        TBContas.Close
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If TBContas!status <> "TÍTULO EM ABERTO" And TBContas!status <> "BLOQUEADA" And TBContas!status <> "TÍTULO LIQUIDADO ANTECIPADO" Then
                        USMsgBox ("Não é permitido excluir esta conta, pois a mesma já foi baixada parcial."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                    If TBContas!Antecipacao = True Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_antecipacao where ID_antecipacao = " & TBContas!IDintconta & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            USMsgBox ("Não é permitido excluir esta conta, pois a mesma é uma antecipação e já esta relacionada a uma conta baixada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            TBAbrir.Close
                            Exit Sub
                        End If
                        TBAbrir.Close
                    End If
                    If TBContas!Despesas_NF = True Then
                        USMsgBox ("Não é permitido excluir esta conta, pois a mesma é uma despesa de importação."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                ElseIf Cmb_opcao_lista = "Status" Then
                        If TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                            USMsgBox ("Não é permitido alterar o status desta conta, pois a mesma é uma antecipação líquidada."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    ElseIf Cmb_opcao_lista = "Baixar" Then
                            If TBContas!status = "BLOQUEADA" Or TBContas!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                                USMsgBox ("Não é permitido baixar esta conta, pois a mesma está bloqueada ou é uma antecipação líquidada."), vbExclamation, "CAPRIND v5.0"
                                .ListItems.Item(InitFor).Checked = False
                                Exit Sub
                            End If
'                            If TBContas!status = "TÍTULO EM ABERTO" Then
'                                Set TBAbrir = CreateObject("adodb.recordset")
'                                TBAbrir.Open "Select * from Familia_financeiro where IDConta = " & TBContas!IDintconta & " and TipoConta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
'                                If TBAbrir.EOF = True Then
'                                    If USMsgBox("Esta conta não está amarrada em nenhuma conta contábil, deseja prosseguir assim mesmo?", vbyesno, "CAPRIND v5.0") = vbNo Then
'                                        .ListItems.Item(InitFor).Checked = False
'                                        TBAbrir.Close
'                                        Exit Sub
'                                    End If
'                                End If
'                                TBAbrir.Close
'                            End If
                            'Verifica se a empresa exige centro de custo
                            If (IsNull(TBContas!ID_nota) = True Or TBContas!ID_nota = "" Or TBContas!ID_nota = "0") And (IsNull(TBContas!Txt_pedido) = True Or TBContas!Txt_pedido = "" Or TBContas!Txt_pedido = "0") Then
                                If ProcVerifExigeCC(TBContas!ID_empresa, "ID_financeiro = " & TBContas!IDintconta, True) = False Then
                                    .ListItems.Item(InitFor).Checked = False
                                    Exit Sub
                                End If
                            End If
                End If
            End If
            TBContas.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lst_contas_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lst_contas.ListItems.Count = 0 Then Exit Sub
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & lst_contas.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    ProcLiberaBotao
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = lst_contas.SelectedItem.index
End If
TBContas.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBContas!ID_empresa) = False And TBContas!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBContas!ID_empresa
txtidintconta.Text = TBContas!IDintconta
Txt_data_transacao.Value = IIf(IsNull(TBContas!Data_transacao), Date, Format(TBContas!Data_transacao, "dd/mm/yyyy"))
txtNDocumento = IIf(IsNull(TBContas!txt_ndocumento), "", TBContas!txt_ndocumento)
If IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" Then txtNPedido = TBContas!Txt_pedido
txtDTEmissao.Value = IIf(IsNull(TBContas!Dt_emissao), Date, Format(TBContas!Dt_emissao, "dd/mm/yyyy"))
txtDtpagto.Value = IIf(IsNull(TBContas!dt_Pagamento), Date, Format(TBContas!dt_Pagamento, "dd/mm/yyyy"))
If TBContas!Antecipacao = True Then Chk_antecipacao.Value = 1 Else Chk_antecipacao.Value = 0
If TBContas!Devolucao = True Then Chk_devolucao.Value = 1 Else Chk_devolucao.Value = 0
If TBContas!Agendado = True Then Chk_agendado.Value = 1 Else Chk_agendado.Value = 0
If TBContas!Conta_fixa = True Then chkConta_fixa.Value = 1 Else chkConta_fixa.Value = 0
If IsNull(TBContas!txt_Parcela) = False And TBContas!txt_Parcela <> "" Then txtparcela.Text = TBContas!txt_Parcela

If TBContas!Tipo = "CL" Then
    Cmb_tipo = "Cliente"
    txtIDFornec = TBContas!int_codforn
ElseIf IsNull(TBContas!Tipo) = True Or TBContas!Tipo = "" Or TBContas!Tipo = "FO" Then
        Cmb_tipo = "Fornecedor"
        txtIDFornec = TBContas!int_codforn
    ElseIf TBContas!Tipo = "FU" Then
            Cmb_tipo = "Funcionário"
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select Codigo from Funcionarios where ID = " & TBContas!int_codforn, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                txtIDFornec = TBFornecedor!CODIGO
            End If
        Else
            Cmb_tipo = "Instituição bancária"
            txtIDFornec = TBContas!int_codforn
End If
txtFornec.Text = IIf(IsNull(TBContas!Txt_fornecedor), "", TBContas!Txt_fornecedor)

'Verifica saldo da antecipação
If TBContas!Antecipacao = True Then qt = IIf(IsNull(TBContas!Saldo_antecipacao), 0, TBContas!Saldo_antecipacao) Else qt = IIf(IsNull(TBContas!dbl_valorpagto), 0, TBContas!dbl_valorpagto)
txtValorTotal.Text = Format(qt, "###,##0.00")

txtObs.Text = IIf(IsNull(TBContas!txt_observacoes), "", TBContas!txt_observacoes)
txt_Competencia = IIf(IsNull(TBContas!Competencia), "", TBContas!Competencia)
txtStatus.Text = IIf(IsNull(TBContas!status), "", TBContas!status)
ProcCarregaPedido

Chk_antecipacao.Enabled = True
Chk_devolucao.Enabled = True
Chk_agendado.Enabled = True
chkConta_fixa.Enabled = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_contas_antecipacao where ID_conta = " & txtidintconta & " or ID_antecipacao = " & txtidintconta & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Or txtStatus = "TÍTULO LIQUIDADO ANTECIPADO" Then
    Chk_antecipacao.Enabled = False
    Chk_devolucao.Enabled = False
Else
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_contas_devolucao where ID_conta = " & txtidintconta & " or ID_devolucao = " & txtidintconta & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Chk_antecipacao.Enabled = False
        Chk_devolucao.Enabled = False
    End If
    TBAbrir.Close
End If

Novo_Pagar = False

NomeCampo = "o tipo do documento"
If IsNull(TBContas!Class_conta) = False And TBContas!Class_conta <> "" Then cmbtipo_conta.Text = TBContas!Class_conta
NomeCampo = "a forma da baixa prevista"
If IsNull(TBContas!FormaBaixa) = False And TBContas!FormaBaixa <> "" Then cmb_forma = TBContas!FormaBaixa
NomeCampo = "a instituição bancária prevista"
If IsNull(TBContas!Banco) = False And TBContas!Banco <> "" Then cmbBanco = TBContas!Banco
1:
    ProcCarregaListaPC

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        If NomeCampo = "a instituição bancária prevista" Then
            USMsgBox ("Não foi encontrado a instituição bancária prevista ou a mesma está bloqueada."), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Não foi encontrado " & NomeCampo & " desta conta."), vbExclamation, "CAPRIND v5.0"
        End If
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaPC()
On Error GoTo tratar_erro

If txtStatus <> "TÍTULO LIQUIDADO ANTECIPADO" Then TextoFiltro = "FF.Pago_recebido = 'False'" Else TextoFiltro = "(FF.Pago_recebido = 'True' or FF.Pago_recebido = 'False')"

Lista_PC.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select FF.ID, F.Codigo, F.txt_descricao, FF.Valor, CP.Antecipacao, CP.Saldo_antecipacao from (tbl_ContasPagar CP INNER JOIN Familia_financeiro FF ON FF.IDConta = CP.IdIntConta) INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDConta = " & txtidintconta & " and FF.Tipoconta = 'P' and " & TextoFiltro & " and FF.Deposito_transf = 'False' order by F.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Lista_PC.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            
            'Verifica saldo da antecipação
            'If TBLISTA!Antecipacao = True Then qt = IIf(IsNull(TBLISTA!Saldo_antecipacao), 0, TBLISTA!Saldo_antecipacao) Else qt = IIf(IsNull(TBLISTA!Valor), 0, TBLISTA!Valor)
            '.Item(.Count).SubItems(3) = Format(qt, "###,##0.00")
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), 0, Format(TBLISTA!valor, "###,##0.00"))
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

Private Sub OptAteomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarMes()
On Error GoTo tratar_erro

NomeRel = "Contas_pagar.rpt"
M = FunVerificaMes(TabFiltro.SelectedItem.key)
If OptDomes.Value = True Then ProcConstruirFiltroPadrao "month(dt_Pagamento)= '" & M & "' and Year(dt_pagamento) = '" & cmbAno & "'", "Month ({tbl_ContasPagar.dt_Pagamento}) = " & M & " and Year ({tbl_ContasPagar.dt_Pagamento}) = " & cmbAno, True, True
If OptAteomes.Value = True Then ProcConstruirFiltroPadrao "month(dt_Pagamento)<= '" & M & "' and Year(dt_pagamento) = '" & cmbAno & "'", "Month ({tbl_ContasPagar.dt_Pagamento}) <= " & M & " and Year ({tbl_ContasPagar.dt_Pagamento}) = " & cmbAno, True, True
If TabFiltro.SelectedItem.key = "Vencidas" Then ProcConstruirFiltroPadrao "(dt_Pagamento) < '" & Date & "'", "{tbl_ContasPagar.dt_Pagamento} < Date(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")", True, True
ProcSalvarDadosRel False, False, False, Date, Date
StrSql_Contas_Pagar_AntecTotal = ""
StrSql_Contas_Pagar_DevTotal = ""
ProcCarregaLista (1)
Imprimir = True
frmContas_pagar_localizar.Todas_contas = False
Novo_Pagar = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtFornec_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: ProcLocalizarFornecedor
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDFornec_Change()
On Error GoTo tratar_erro

txtFornec = ""
If txtIDFornec <> "" Then
    If Cmb_tipo <> "Funcionário" Then
        VerifNumero = txtIDFornec
        ProcVerificaNumero
        If VerifNumero = False Then
            txtIDFornec = ""
            txtIDFornec.SetFocus
            Exit Sub
        End If
    End If
    Set TBAbrir = CreateObject("adodb.recordset")
    If Cmb_tipo = "Cliente" Then
        TBAbrir.Open "Select NomeRazao, Banco, Tipo_doc from Clientes where idcliente = " & txtIDFornec & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtFornec = TBAbrir!NomeRazao
            NomeCampo = "instituição bancária prevista do cliente."
            If IsNull(TBAbrir!Banco) = False And TBAbrir!Banco <> "" Then cmbBanco.Text = TBAbrir!Banco
            NomeCampo = "tipo do documento do cliente."
            If IsNull(TBAbrir!Tipo_doc) = False And TBAbrir!Tipo_doc <> "" Then cmbtipo_conta.Text = TBAbrir!Tipo_doc
        End If
    ElseIf Cmb_tipo = "Fornecedor" Then
            TBAbrir.Open "Select Nome_Razao, Banco, Tipo_doc from compras_fornecedores where idcliente = " & txtIDFornec & " and Prospecto = 'False' and DtValidacao IS NOT NULL and status <> 'Bloqueado'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                txtFornec = TBAbrir!Nome_Razao
                NomeCampo = "instituição bancária prevista do fornecedor."
                If IsNull(TBAbrir!Banco) = False And TBAbrir!Banco <> "" Then cmbBanco.Text = TBAbrir!Banco
                NomeCampo = "tipo do documento do fornecedor."
                If IsNull(TBAbrir!Tipo_doc) = False And TBAbrir!Tipo_doc <> "" Then cmbtipo_conta.Text = TBAbrir!Tipo_doc
            End If
        ElseIf Cmb_tipo = "Funcionário" Then
                TBAbrir.Open "Select Nome from Funcionarios where Codigo = '" & txtIDFornec & "' and DtValidacao IS NOT NULL and Situacao = 'Normal'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtFornec = TBAbrir!Nome
            Else
                TBAbrir.Open "Select Txt_descricao from tbl_Instituicoes where ID = " & txtIDFornec & " and DtValidacao IS NOT NULL and Bloqueado <> 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then txtFornec = TBAbrir!Txt_descricao
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtIDFornec_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: ProcLocalizarFornecedor
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

Private Sub txtValorTotal_Change()
On Error GoTo tratar_erro
    
If txtValorTotal.Text <> "" Then
    VerifNumero = txtValorTotal.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtValorTotal.Text = ""
        txtValorTotal.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValorTotal_LostFocus()
On Error GoTo tratar_erro
    
txtValorTotal.Text = Format(txtValorTotal.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaPedido()
On Error GoTo tratar_erro

With txtNPedido
    .Clear
    
    If IsNull(TBContas!ID_nota) = False And TBContas!ID_nota <> "" And TBContas!ID_nota <> "0" Then
        Set TBProposta = CreateObject("adodb.recordset")
        TBProposta.Open "Select PN.Proposta from tbl_proposta_nota PN INNER JOIN tbl_Dados_Nota_Fiscal NF ON PN.ID_nota = NF.ID where PN.ID_nota = " & TBContas!ID_nota & " and NF.int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
        If TBProposta.EOF = False Then
            Do While TBProposta.EOF = False
                If IsNull(TBProposta!Proposta) = False Then .AddItem TBProposta!Proposta
                TBProposta.MoveNext
            Loop
        Else
            If IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" And TBContas!Txt_pedido <> "0" Then
                .AddItem TBContas!Txt_pedido
                .Text = TBContas!Txt_pedido
            End If
        End If
        TBProposta.Close
    Else
        If IsNull(TBContas!Txt_pedido) = False And TBContas!Txt_pedido <> "" And TBContas!Txt_pedido <> "0" Then
            .AddItem TBContas!Txt_pedido
            .Text = TBContas!Txt_pedido
        End If
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRelacionamento()
On Error GoTo tratar_erro

If txtidintconta = "" Then
    USMsgBox ("Informe a conta antes de visualizar a lista de contas relacionadas."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Financeiro_Contas_Pagas = False
Financeiro_Contas_Pagar = True
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
frmContas_antecipacoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLiberaBotao()
On Error GoTo tratar_erro

With USToolBar1
    If TBContas!Antecipacao = True Then .ButtonState(13) = 0 Else .ButtonState(13) = 5
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvarDadosRel(DataTransacao As Boolean, DataEmissao As Boolean, DataVencimento As Boolean, DataInicio As Date, DataFinal As Date)
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatoriosTotal
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    TBLISTA!Responsavel = pubUsuario
    TBLISTA!Modulo = Formulario
    If DataTransacao = True Or DataEmissao = True Or DataVencimento = True Then
        TBLISTA!Data_inicial = DataInicio
        TBLISTA!Data_final = DataFinal
        TBLISTA!Turno = True
    Else
        TBLISTA!Turno = False
    End If
    TBLISTA.Update
End If
TBLISTA.Close

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
    Case 6: ProcPlanoContas
    Case 7: ProcAgendaDia
    Case 8: ProcParcelar
    Case 9: ProcCopiar
    Case 10: ProcPagar
    Case 11: ProcStatus
    Case 12: ProcCC
    Case 13: ProcRelacionamento
    Case 14: ProcAgendarPgto
    Case 16: ProcAjuda
    Case 17: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifExigeCC(ID_empresa As Integer, TextoFiltro As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifExigeCC = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and CC_obrigatorio = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select ID from CC_Realizado where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = True Then
        If MostrarMsg = True Then USMsgBox ("É obrigatório cadastrar o(s) centro(s) de custo para esta conta."), vbInformation, "CAPRIND v5.0"
        ProcVerifExigeCC = False
    End If
    TBFI.Close
End If
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
