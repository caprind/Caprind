VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestoque_item 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Movimentação"
   ClientHeight    =   10035
   ClientLeft      =   3870
   ClientTop       =   2325
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView Lista 
      Height          =   2085
      Left            =   60
      TabIndex        =   4
      Top             =   1800
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   3678
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Lote"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. int."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cod. ref."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   11113
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   707
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Famíllia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Local de armaz."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Corrida"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Certificado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "N. série"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "Cliente/fornecedor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Qtde PÇ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "Vlr. unit."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "Vlr. total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Object.Tag             =   "T"
         Text            =   "Liberado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   18
         Text            =   "Part Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Fabricante"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView Lista_movimentacao 
      Height          =   3015
      Left            =   60
      TabIndex        =   5
      Top             =   5940
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   5318
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
      Appearance      =   0
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
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Operação"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Entrada"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Entrada PÇ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Saída"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Saída PÇ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Documento"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Requisitante"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "Destino"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "T"
         Text            =   "PC/PI"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Tag             =   "T"
         Text            =   "Fornecedor/Cliente"
         Object.Width           =   2385
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Observações"
         Object.Width           =   2385
      EndProperty
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registros de RE's"
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
      Left            =   55
      TabIndex        =   32
      Top             =   5310
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
         Left            =   6060
         TabIndex        =   41
         Text            =   "10"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtPagIr 
         Height          =   315
         Left            =   10140
         TabIndex        =   33
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   12360
         TabIndex        =   34
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":0000
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
         Left            =   11820
         TabIndex        =   35
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":37A4
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
         Left            =   10710
         TabIndex        =   36
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
         Left            =   11280
         TabIndex        =   37
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":72AD
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
         Left            =   12900
         TabIndex        =   38
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":B39C
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
      Begin VB.Label Label23 
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
         Left            =   6690
         TabIndex        =   61
         Top             =   240
         Width           =   1440
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
         Left            =   5370
         TabIndex        =   42
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
         Left            =   450
         TabIndex        =   40
         Top             =   300
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
         Left            =   13650
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   31
      Top             =   9780
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações para pesquisa"
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
      Height          =   855
      Index           =   0
      Left            =   55
      TabIndex        =   25
      Top             =   960
      Width           =   15195
      Begin DrawSuite2022.USCheckBox chkEstoquePositivo 
         Height          =   225
         Left            =   2850
         TabIndex        =   71
         Top             =   450
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   397
         Caption         =   "Com saldo no estoque"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   0   'False
         Value           =   1
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
         ItemData        =   "frmEstoque_item.frx":EC28
         Left            =   150
         List            =   "frmEstoque_item.frx":EC2A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   2625
      End
      Begin VB.OptionButton optIgual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Igual"
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
         Left            =   14310
         TabIndex        =   65
         Top             =   120
         Width           =   705
      End
      Begin VB.OptionButton Optmeio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meio"
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
         Left            =   13050
         TabIndex        =   64
         Top             =   120
         Width           =   705
      End
      Begin VB.OptionButton Optinicio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Início"
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
         Left            =   12270
         TabIndex        =   63
         Top             =   120
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Optfim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fim"
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
         Left            =   13740
         TabIndex        =   62
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   12060
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   3015
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmEstoque_item.frx":EC2C
         Left            =   10050
         List            =   "frmEstoque_item.frx":EC5A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   1995
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frmEstoque_item.frx":ED0A
         Left            =   12060
         List            =   "frmEstoque_item.frx":ED0C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   3015
      End
      Begin DrawSuite2022.USCheckBox chkTerceiros 
         Height          =   225
         Left            =   4890
         TabIndex        =   72
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         Caption         =   "Em terceiros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USCheckBox Chk_com_empenho 
         Height          =   225
         Left            =   6120
         TabIndex        =   73
         Top             =   450
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   397
         Caption         =   "Empenhado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         Value           =   1
      End
      Begin DrawSuite2022.USCheckBox Chk_sem_empenho 
         Height          =   225
         Left            =   7380
         TabIndex        =   74
         Top             =   450
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   397
         Caption         =   "Sem empenho"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         Value           =   1
      End
      Begin DrawSuite2022.USCheckBox chkBloqueados 
         Height          =   225
         Left            =   8850
         TabIndex        =   75
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         Caption         =   "Bloqueado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   10695
         TabIndex        =   28
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00000080&
      Height          =   1425
      Left            =   55
      TabIndex        =   26
      Top             =   3885
      Width           =   15195
      Begin VB.TextBox txtLocalArmazenamento 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Total de peça em estoque do RE."
         Top             =   1005
         Width           =   5115
      End
      Begin VB.TextBox txtLote 
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
         Left            =   90
         TabIndex        =   80
         TabStop         =   0   'False
         ToolTipText     =   "Número de série."
         Top             =   435
         Width           =   885
      End
      Begin VB.TextBox txtTTEntrada 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Height          =   315
         Left            =   8940
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         ToolTipText     =   "Total entrada do item"
         Top             =   435
         Width           =   1245
      End
      Begin VB.TextBox txtTTSaida 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
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
         Height          =   315
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "Total saida do item"
         Top             =   435
         Width           =   1125
      End
      Begin VB.TextBox Txt_n_serie 
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
         Left            =   3060
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Número de série."
         Top             =   435
         Width           =   1065
      End
      Begin VB.TextBox Txt_cod_ref 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2130
         TabIndex        =   7
         ToolTipText     =   "Código de referência."
         Top             =   435
         Width           =   915
      End
      Begin VB.TextBox txtlocalização 
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
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Registro de entrada."
         Top             =   435
         Width           =   765
      End
      Begin VB.TextBox Txt_qtde_est_dispRE 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Height          =   315
         Left            =   13875
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Estoque disponível."
         Top             =   435
         Width           =   1170
      End
      Begin VB.TextBox Txt_qtde_empenhoRE 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Height          =   315
         Left            =   12375
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Total empenhado."
         Top             =   435
         Width           =   1125
      End
      Begin VB.TextBox Txt_valor_total_estRE 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   7890
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Valor total em estoque."
         Top             =   435
         Width           =   1035
      End
      Begin VB.TextBox Txt_valor_unitRE 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Custo médio unitário."
         Top             =   435
         Width           =   1065
      End
      Begin VB.TextBox txtFornecedor 
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
         Left            =   4290
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Estoque de peça disponível."
         Top             =   1005
         Width           =   5655
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Total de peça em estoque do RE."
         Top             =   1005
         Width           =   4185
      End
      Begin VB.TextBox Txt_qtde_est_tercRE 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   6060
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Total em terceiros do RE."
         Top             =   435
         Width           =   735
      End
      Begin VB.TextBox Txt_qtde_estoqueRE 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Height          =   315
         Left            =   11340
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Total em estoque do RE."
         Top             =   435
         Width           =   1020
      End
      Begin DrawSuite2022.USButton Cmd_empenhoRE 
         Height          =   315
         Left            =   13530
         TabIndex        =   66
         ToolTipText     =   "Buscar empenhos da RE"
         Top             =   435
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":ED0E
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
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton Cmd_salvar_ref_serie 
         Height          =   315
         Left            =   4140
         TabIndex        =   68
         ToolTipText     =   "Salvar numero de série"
         Top             =   435
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":2CE13
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin DrawSuite2022.USButton btnLote 
         Height          =   315
         Left            =   990
         TabIndex        =   81
         ToolTipText     =   "Salvar numero de série"
         Top             =   430
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":35818
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin DrawSuite2022.USButton btnSalvarVencimento 
         Height          =   315
         Left            =   5715
         TabIndex        =   85
         ToolTipText     =   "Salvar data de vencimento do lote"
         Top             =   435
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":3E21D
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin MSComCtl2.DTPicker txtvencimento 
         Height          =   315
         Left            =   4470
         TabIndex        =   87
         ToolTipText     =   "Data do vencimento do lote."
         Top             =   435
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   197263361
         CurrentDate     =   39057
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
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
         Left            =   4470
         TabIndex        =   86
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   11610
         TabIndex        =   84
         Top             =   810
         Width           =   1815
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
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
         Left            =   420
         TabIndex        =   82
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total saida (=)"
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
         Left            =   10215
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total entrada (-)"
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
         Left            =   8955
         TabIndex        =   78
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "   Informações do estoque do RE (Saldos em estoque)  "
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
         Left            =   90
         TabIndex        =   69
         Top             =   0
         Width           =   3960
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
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
         TabIndex        =   60
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero série"
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
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RE"
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
         Left            =   1590
         TabIndex        =   58
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Empenhado"
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
         Left            =   12390
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Disponível"
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
         Left            =   13890
         TabIndex        =   56
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total"
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
         Left            =   8070
         TabIndex        =   55
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor unitário"
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
         Left            =   6855
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor | Cliente"
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
         Left            =   6375
         TabIndex        =   53
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
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
         Left            =   510
         TabIndex        =   43
         Top             =   810
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro"
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
         Left            =   6135
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em estoque"
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
         Left            =   11385
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   29
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   17
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
      ButtonCaption2  =   "Entrada"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Entrada."
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
      ButtonWidth2    =   46
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Retirada"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Retirada."
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
      ButtonLeft3     =   88
      ButtonTop3      =   2
      ButtonWidth3    =   49
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Excluir"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Excluir movimentação (F4)"
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
      ButtonLeft4     =   139
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
      ButtonLeft5     =   180
      ButtonTop5      =   2
      ButtonWidth5    =   51
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Status"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Status."
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
      ButtonLeft6     =   233
      ButtonTop6      =   2
      ButtonWidth6    =   39
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Estrutura"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Estrutura (F7)"
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
      ButtonLeft7     =   274
      ButtonTop7      =   2
      ButtonWidth7    =   53
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Sucata/Retalho"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Gerar sucata/retalho (F8)"
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
      ButtonLeft8     =   329
      ButtonTop8      =   2
      ButtonWidth8    =   82
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonCaption9  =   "Excluir sucata/ret."
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonToolTipText9=   "Excluir entrada de sucata/retalho (F9)"
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
      ButtonLeft9     =   413
      ButtonTop9      =   2
      ButtonWidth9    =   96
      ButtonHeight9   =   21
      ButtonUseMaskColor9=   0   'False
      ButtonCaption10 =   "Local armaz."
      ButtonEnabled10 =   0   'False
      ButtonIconSize10=   32
      ButtonToolTipText10=   "Alterar local de armazenamento do produto (F10)"
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
      ButtonLeft10    =   511
      ButtonTop10     =   2
      ButtonWidth10   =   68
      ButtonHeight10  =   21
      ButtonUseMaskColor10=   0   'False
      ButtonCaption11 =   "Valor unit."
      ButtonEnabled11 =   0   'False
      ButtonIconSize11=   32
      ButtonToolTipText11=   "Alterar valor unitário do RE (F11)"
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
      ButtonLeft11    =   581
      ButtonTop11     =   2
      ButtonWidth11   =   57
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft12    =   640
      ButtonTop12     =   2
      ButtonWidth12   =   85
      ButtonHeight12  =   21
      ButtonUseMaskColor12=   0   'False
      ButtonCaption13 =   "Atualizar"
      ButtonEnabled13 =   0   'False
      ButtonIconSize13=   32
      ButtonToolTipText13=   "Utilizado pelo administrador do sistema."
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
      ButtonLeft13    =   727
      ButtonTop13     =   2
      ButtonWidth13   =   50
      ButtonHeight13  =   21
      ButtonUseMaskColor13=   0   'False
      ButtonEnabled14 =   0   'False
      ButtonIconSize14=   32
      ButtonAlignment14=   2
      ButtonType14    =   1
      ButtonStyle14   =   -1
      BeginProperty ButtonFont14 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState14   =   -1
      ButtonLeft14    =   779
      ButtonTop14     =   4
      ButtonWidth14   =   2
      ButtonHeight14  =   54
      ButtonCaption15 =   "Ajuda"
      ButtonEnabled15 =   0   'False
      ButtonIconSize15=   32
      ButtonToolTipText15=   "Ajuda (F1)"
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
      ButtonLeft15    =   783
      ButtonTop15     =   2
      ButtonWidth15   =   36
      ButtonHeight15  =   21
      ButtonUseMaskColor15=   0   'False
      ButtonCaption16 =   "Sair"
      ButtonEnabled16 =   0   'False
      ButtonIconSize16=   32
      ButtonToolTipText16=   "Sair (Esc)"
      ButtonKey16     =   "16"
      ButtonAlignment16=   2
      BeginProperty ButtonFont16 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft16    =   821
      ButtonTop16     =   2
      ButtonWidth16   =   26
      ButtonHeight16  =   21
      ButtonUseMaskColor16=   0   'False
      ButtonEnabled17 =   0   'False
      ButtonIconSize17=   32
      ButtonKey17     =   "17"
      BeginProperty ButtonFont17 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState17   =   5
      ButtonLeft17    =   849
      ButtonTop17     =   2
      ButtonWidth17   =   24
      ButtonHeight17  =   24
      ButtonUseMaskColor17=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13020
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_item.frx":46C22
         Count           =   1
      End
   End
   Begin VB.Frame Frame5 
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
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   60
      TabIndex        =   44
      Top             =   8940
      Width           =   15195
      Begin VB.TextBox Txt_qtde_est_disp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Height          =   315
         Left            =   13575
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Estoque disponível."
         Top             =   435
         Width           =   1500
      End
      Begin VB.TextBox Txt_qtde_empenho 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Height          =   315
         Left            =   11895
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Total empenhado."
         Top             =   435
         Width           =   1275
      End
      Begin VB.TextBox Txt_qtde_estoque 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Height          =   315
         Left            =   10290
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Total em estoque."
         Top             =   435
         Width           =   1575
      End
      Begin VB.TextBox Txt_qtde_est_terc 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Total em terceiros."
         Top             =   430
         Width           =   1410
      End
      Begin VB.TextBox Txt_valor_total_est 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Valor total em estoque."
         Top             =   430
         Width           =   1665
      End
      Begin VB.TextBox Txt_custo_medio_unit 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Custo médio unitário."
         Top             =   430
         Width           =   1575
      End
      Begin VB.TextBox Txt_qtde_estoque_PC 
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
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Total de peça em estoque."
         Top             =   1845
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox Txt_qtde_est_disp_PC 
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
         Left            =   8025
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Estoque de peça disponível."
         Top             =   1845
         Visible         =   0   'False
         Width           =   1725
      End
      Begin DrawSuite2022.USButton Cmd_empenho 
         Height          =   315
         Left            =   13200
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Buscar empenhos do item"
         Top             =   435
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_item.frx":507F2
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
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "   Informações do estoque do Iem (Saldo em estoque)  "
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
         Left            =   90
         TabIndex        =   70
         Top             =   0
         Width           =   3960
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo em estoque"
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
         TabIndex        =   52
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Empenhado"
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
         Left            =   11985
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(=) Disponível"
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
         Left            =   13740
         TabIndex        =   50
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total em terceiros"
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
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total"
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
         Left            =   3660
         TabIndex        =   48
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ct. médio unitário"
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
         Left            =   1740
         TabIndex        =   47
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est. disponível PÇ"
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
         Left            =   8145
         TabIndex        =   46
         Top             =   1650
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total em est. PÇ"
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
         Left            =   2610
         TabIndex        =   45
         Top             =   1650
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmestoque_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Documento_Ordem As String 'OK
Dim Sql_Estoque_Movimentacao As String 'OK
Dim TBLISTA_Estoque_Movimentacao As ADODB.Recordset 'OK
Dim Status_movimentacao As String 'OK

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

txtlocalização = ""
Txt_cod_ref = ""
Txt_n_serie = ""
Txt_qtde_estoqueRE = "0,0000"
Txt_qtde_estoque_PCRE = "0,0000"
Txt_qtde_empenhoRE = "0,0000"
Txt_qtde_est_dispRE = "0,0000"
Txt_qtde_est_disp_PCRE = "0,0000"
Txt_qtde_est_tercRE = "0,0000"
Txt_valor_total_estRE = "0,00"
Txt_valor_unitRE = "0,0000"

Txt_qtde_estoque = "0,0000"
Txt_qtde_estoque_PC = "0,0000"
Txt_qtde_empenho = "0,0000"
'Txt_qtde_est_disp = "0,0000"
Txt_qtde_est_disp_PC = "0,0000"
Txt_qtde_est_terc = "0,0000"
'Txt_valor_total_est = "0,00"
Txt_custo_medio_unit = "0,00000"

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
Lista_Movimentacao.ListItems.Clear
If Sql_Estoque_Movimentacao = "" Then Exit Sub
Set TBLISTA_Estoque_Movimentacao = CreateObject("adodb.recordset")
'Debug.print Sql_Estoque_Movimentacao
TBLISTA_Estoque_Movimentacao.Open Sql_Estoque_Movimentacao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Estoque_Movimentacao.EOF = False Then ProcExibePagina (Pagina) Else CodigoLista = 0
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
TBLISTA_Estoque_Movimentacao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Estoque_Movimentacao.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Estoque_Movimentacao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Estoque_Movimentacao.RecordCount - IIf(Pagina > 1, (TBLISTA_Estoque_Movimentacao.PageSize * (Pagina - 1)), 0), TBLISTA_Estoque_Movimentacao.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Estoque_Movimentacao.EOF = False And (ContadorReg <= TamanhoPagina)

'===============================================================================
' Busca saldo total da RE
'===============================================================================
StrSql = "select Sum(Entrada) - Sum(Saida) as saldo from Estoque_movimentacao where IdEstoque = '" & TBLISTA_Estoque_Movimentacao!IDEstoque & "'"
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
   vlr_Saldo = IIf(IsNull(TBFIltro!Saldo), "", Format(TBFIltro!Saldo, "###,##0.0000"))
   If TBFIltro!Saldo <> "" Then
   ValorTotal = TBFIltro!Saldo * TBLISTA_Estoque_Movimentacao!valor_unitario
   End If
End If
TBFIltro.Close
'===============================================================================

Cliente = ""
If IsNull(TBLISTA_Estoque_Movimentacao!Fornecedor) = False And TBLISTA_Estoque_Movimentacao!Fornecedor <> "" Then
    Cliente = TBLISTA_Estoque_Movimentacao!Fornecedor
ElseIf IsNull(TBLISTA_Estoque_Movimentacao!Cliente) = False And TBLISTA_Estoque_Movimentacao!Cliente <> "" Then
    Cliente = TBLISTA_Estoque_Movimentacao!Cliente
End If

If chkEstoquePositivo.Value = 1 And vlr_Saldo = 0 Then GoTo Proximo
    With Lista.ListItems
        .Add , , IIf(IsNull(TBLISTA_Estoque_Movimentacao!IDEstoque), 0, TBLISTA_Estoque_Movimentacao!IDEstoque)
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Data), "", Format(TBLISTA_Estoque_Movimentacao!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!LOTE), "", TBLISTA_Estoque_Movimentacao!LOTE)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Desenho), "", TBLISTA_Estoque_Movimentacao!Desenho)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Ref), "", TBLISTA_Estoque_Movimentacao!Ref)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Descricao), "", TBLISTA_Estoque_Movimentacao!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Unidade), "", TBLISTA_Estoque_Movimentacao!Unidade)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Classe), "", TBLISTA_Estoque_Movimentacao!Classe)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!local_armaz), "", TBLISTA_Estoque_Movimentacao!local_armaz)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Corrida), "", TBLISTA_Estoque_Movimentacao!Corrida)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Certificado), "", TBLISTA_Estoque_Movimentacao!Certificado)
        .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!Numero_serie), "", TBLISTA_Estoque_Movimentacao!Numero_serie)
        .Item(.Count).SubItems(12) = Cliente
        .Item(.Count).SubItems(13) = Format(vlr_Saldo, "###,##0.0000")
        .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA_Estoque_Movimentacao!valor_unitario), "", Format(TBLISTA_Estoque_Movimentacao!valor_unitario, "###,##0.0000"))
        .Item(.Count).SubItems(16) = Format(ValorTotal, "###,##0.0000")
        .Item(.Count).SubItems(17) = TBLISTA_Estoque_Movimentacao!Liberado '"Sim"
        
        Set TBItem = CreateObject("adodb.recordset")
        StrSql = "Select Part_number, Fabricante from Projproduto_fabricante PF Inner join Fabricante_Marca FM on PF.IDFabricante = FM.Id where codproduto = " & TBLISTA_Estoque_Movimentacao!Codproduto
        'Debug.print StrSql
        
        TBItem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            .Item(.Count).SubItems(18) = IIf(IsNull(TBItem!Part_number), "", (TBItem!Part_number))
            .Item(.Count).SubItems(19) = IIf(IsNull(TBItem!Fabricante), "", (TBItem!Fabricante))
        Else
            .Item(.Count).SubItems(18) = ""
            .Item(.Count).SubItems(19) = ""
        End If
        
        
    End With
Proximo:
    
TBLISTA_Estoque_Movimentacao.MoveNext
ContadorReg = ContadorReg + 1
Contador = Contador + 1
PBLista.Value = Contador
ValorTotalEstoque = ValorTotalEstoque + ValorTotal
'TotalEmEstoque = TotalEmEstoque + vlr_Saldo
Loop

lblRegistros.Caption = "Nº de registros: " & TBLISTA_Estoque_Movimentacao.RecordCount
If TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Estoque_Movimentacao.PageCount
ElseIf TBLISTA_Estoque_Movimentacao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.PageCount & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Estoque_Movimentacao.AbsolutePage - 1 & " de: " & TBLISTA_Estoque_Movimentacao.PageCount
End If

Txt_valor_total_est = Format(ValorTotalEstoque, "###,##0.000000")
'Txt_qtde_est_disp = Format(TotalEmEstoque, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnLote_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente alterar o numero do lote nessa movimentação?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute ("Update Estoque_controle SET Lote = '" & txtLote.Text & "' where idestoque = '" & txtlocalização.Text & "'")
    Conexao.Execute ("Update Estoque_movimentacao SET Lote = '" & txtLote.Text & "' where idestoque = '" & txtlocalização.Text & "'")
    USMsgBox "Lote alterado com sucesso!", vbInformation, "CAPRIND v5.0"
    ProcAtualizalista (1)
    ProcCarregaListaMovimentacao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnSalvarVencimento_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If txtvencimento = "" Then
    NomeCampo = "o vencimento"
    ProcVerificaAcao
    txtvencimento.SetFocus
    Exit Sub
End If

If USMsgBox("Deseja realmente alterar o vencimento do lote deste RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    
    Conexao.Execute "Update estoque_controle Set Vencimento = '" & txtvencimento & "' where IDestoque = " & txtlocalização
    Conexao.Execute "Update Estoque_fisico Set Vencimento = '" & txtvencimento & "' where IDestoque = " & txtlocalização
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Movimentação"
    Evento = "Alterar vencimento"
    ID_documento = txtlocalização
    Documento = "Cód. interno: " & Lista.SelectedItem.SubItems(3)
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Lista_Movimentacao.ListItems.Clear
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub chkBloqueados_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkEstoquePositivo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkTerceiros_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_com_empenho_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_sem_empenho_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais
IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

txtTexto.Visible = True
cmbTexto.Visible = False
If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Grupo" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    cmbTexto.Clear
    
    Select Case cmbfiltrarpor
        Case "Local de armazenamento": ProcCarregaComboLA cmbTexto, True, True
        Case "Família": ProcCarregaComboFamilia cmbTexto, "Familia is not null", False
    End Select
ElseIf cmbfiltrarpor = "RE" And txtTexto <> "" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlterar_valor()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista.ListItems.Count = 0 Then Exit Sub
If USMsgBox("Deseja realmente atualizar o valor unitário deste RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem1:
     Valorunitario = InputBox("Favor informar o valor unitário.")
     If Valorunitario = "" Then Exit Sub
     If IsNumeric(Valorunitario) = False Then
         USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
         GoTo Mensagem1
     End If
     If Valorunitario = 0 Then
         USMsgBox ("Não é permitido colocar zero no valor unitário."), vbExclamation, "CAPRIND v5.0"
         GoTo Mensagem1
     End If
     valor = Valorunitario
     NovoValor = Replace(valor, ",", ".")
     Conexao.Execute "UPDATE Estoque_Controle Set valor_unitario = " & NovoValor & " where IDEstoque = " & Lista.SelectedItem
     Conexao.Execute "UPDATE Estoque_Controle Set Valor_total = ROUND(valor_unitario * Estoque_real, 2) where IDEstoque = " & Lista.SelectedItem
     Conexao.Execute "UPDATE Estoque_movimentacao Set VlrUnit = " & NovoValor & " where IDEstoque = " & Lista.SelectedItem
     Conexao.Execute "UPDATE Estoque_movimentacao Set VlrTotal = ROUND(VlrUnit * Entrada, 2) where IDEstoque = " & Lista.SelectedItem & " and Entrada <> 0"
     Conexao.Execute "UPDATE Estoque_movimentacao Set VlrTotal = ROUND(VlrUnit * Saida, 2) where IDEstoque = " & Lista.SelectedItem & " and Saida <> 0"
     Conexao.Execute "Update CC set CC.Valor = EM.VlrTotal from CC_realizado CC INNER JOIN Estoque_movimentacao EM on CC.ID_estoque = EM.Idoperacao where ID_estoque = " & Lista.SelectedItem
    
     Set TBFIltro = CreateObject("adodb.recordset")
     TBFIltro.Open "Select Documento from Estoque_movimentacao where IDEstoque = " & Lista.SelectedItem & " and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL') group by Documento", Conexao, adOpenKeyset, adLockOptimistic
     If TBFIltro.EOF = False Then
         Do While TBFIltro.EOF = False
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select IDlista from Compras_pedido_lista where Ordem = " & TBFIltro!Documento & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then
                ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBFIltro!Documento
            End If
            TBAbrir.Close
             TBFIltro.MoveNext
         Loop
     End If
     TBFIltro.Close
     USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
     '==================================
     Modulo = "Estoque/Movimentação"
     Evento = "Alterar valor unitário do inventário"
     ID_documento = txtlocalização
     Documento = "Cód. interno: " & Lista.SelectedItem.ListSubItems(3)
     Documento1 = ""
     ProcGravaEvento
     '==================================
     ProcFiltrar
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Conexao.Execute ("update Estoque_movimentacao Set Bloqueado = EC.Bloqueado from Estoque_movimentacao EM inner Join Estoque_controle EC On EM.idestoque = EC.idEstoque")
Conexao.Execute ("update Estoque_movimentacao Set ID_empresa = EC.id_empresa from Estoque_movimentacao EM inner Join Estoque_controle EC On EM.idestoque = EC.idEstoque")
Conexao.Execute ("update Estoque_movimentacao Set Familia = PP.Classe, Unidade = PP.Unidade from Estoque_movimentacao EM inner Join ProjProduto PP On EM.Desenho = PP.Desenho")
Conexao.Execute ("update Estoque_movimentacao Set Grupo = PF.Grupo from Estoque_movimentacao EM inner Join Projfamilia PF On EM.Familia = PF.Familia")


Acao = "filtrar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If Chk_com_empenho.Value = 0 And Chk_sem_empenho.Value = 0 Then
    NomeCampo = "uma das opções com ou sem empenho"
    ProcVerificaAcao
    Exit Sub
End If
'If chkEstoquePositivo.Value = 1 Then TextoFiltroQtde = " and Saldo > 0" Else TextoFiltroQtde = ""
If chkEstoquePositivo.Value = 1 Then TextoFiltroQtde = " and Estoque_disponivel > 0" Else TextoFiltroQtde = ""
'If chkTerceiros.Value = 1 Then TextoFiltroTerc = "and EP.destino = 'Terceiros' and EP.Terceiros = 'True'" Else TextoFiltroTerc = ""
'If Chk_com_empenho.Value = 1 And Chk_sem_empenho.Value = 1 Then
'    TextoFiltroEmp = "and EP.Qtde_empenhada >= 0"
'ElseIf Chk_com_empenho.Value = 1 Then
'        TextoFiltroEmp = "and EP.Qtde_empenhada > 0"
'    Else
'        TextoFiltroEmp = "and EP.Qtde_empenhada = 0"
'End If
If chkBloqueados.Value = 1 Then StatusFiltro = "and EP.Bloqueado = 'True'" Else StatusFiltro = " and EP.Bloqueado = 'False'"

'CamposFiltro = "EP.idestoque, EP.Etiqueta, EP.Data, EP.LOTE, EP.Codigo, EP.Ref, EP.Descricao, EP.classe, EP.local_armaz, EP.Corrida, EP.Certificado, EP.Numero_serie, EP.Fornecedor, EP.Cliente, EP.Un, EP.valor_unitario, EP.Valor_Total, EP.Liberado"
'INNERJOINTEXTO = "Select " & CamposFiltro & " from (Estoque_controle_Saldo_RE EP LEFT JOIN estoque_movimentacao EM ON EM.IDestoque = EP.IDestoque) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = EP.codproduto"

CamposFiltro = "EP.codProduto, EP.idestoque, EP.Etiqueta, EP.Data, EP.LOTE, EP.Desenho, EP.Ref, EP.Descricao, EP.classe, EP.local_armaz, EP.Corrida, EP.Certificado, EP.Numero_serie, EP.Fornecedor, EP.Cliente, EP.Unidade, EP.valor_unitario, EP.Valor_Total, EP.Liberado"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (Estoque_Produtos EP LEFT JOIN estoque_movimentacao EM ON EM.IDestoque = EP.IDestoque) left outer join Projproduto_fabricante PFAB On PFAB.Codproduto = EP.codproduto "


'CamposFiltro = "idestoque, Etiqueta, Data, LOTE, Codigo, Ref, Descricao, classe, local_armaz,Corrida, Certificado, Numero_serie, Fornecedor, Cliente, Un, valor_unitario, Valor_Total, Bloqueado"
'INNERJOINTEXTO = "Select " & CamposFiltro & " from Estoque_controle_Saldo_RE "
TextoFiltroPadrao = "(EP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " or EP.ID_empresa is null) " & TextoFiltroQtde & TextoFiltroTerc & TextoFiltroEmp & StatusFiltro & " group by " & CamposFiltro & " order by Desenho, idestoque"

If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
    Select Case cmbfiltrarpor
        Case "Local de armazenamento": TextoFiltro = "EP.local_armaz"
        Case "Família": TextoFiltro = "EP.classe"
        Case "Grupo": TextoFiltro = "EP.Grupo"
        Case "Código interno": TextoFiltro = "EP.Desenho"
        Case "Código de referência": TextoFiltro = "EP.Ref"
        Case "Corrida": TextoFiltro = "EP.Corrida"
        Case "Certificado": TextoFiltro = "EP.Certificado"
        Case "Descrição": TextoFiltro = "EP.descricao"
        Case "Lote": TextoFiltro = "EP.Lote"
        Case "Etiqueta": TextoFiltro = "EP.Etiqueta"
        Case "Documento": TextoFiltro = "EM.Documento"
        Case "Número de série": TextoFiltro = "EP.Numero_serie"
        Case "Part number": TextoFiltro = "PFAB.Part_number"
        Case "RE": TextoFiltro = "EM.IDestoque"
    End Select
        
    
    If cmbfiltrarpor = "Local de armazenamento" Or cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Grupo" Then
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
    Else
        Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    Sql_Estoque_Movimentacao = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

'Debug.print Sql_Estoque_Movimentacao

ProcAtualizalista (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_atualizar_qtde_est_RE_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If Lista.SelectedItem = "" Or Lista.SelectedItem = "0" Then Exit Sub
If USMsgBox("Deseja realmente atualizar a quantidade em estoque do RE " & Lista.SelectedItem & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select * from estoque_controle where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from projproduto where Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If TBAbrir!Estoque = True Then
                qtdeliberada = 0
                qtdeliberadaPC = 0
                qtdeliberar = 0
                qtdeliberarPC = 0
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Sum(EM.Entrada) as qtdeliberada, Sum(ISNULL(EM.Entrada_PC, 0)) as qtdeliberadaPC, Sum(EM.Saida) as qtdeliberar, Sum(ISNULL(EM.Saida_PC, 0)) as qtdeliberarPC from ((Estoque_movimentacao EM LEFT JOIN Estoque_controle_recebimento ECR ON ECR.Id = EM.IDEstoque_recebimento) LEFT JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido AND CPL.IdLista = ECR.IdLista) LEFT JOIN tbl_NaturezaOperacao CFOP ON CPL.ID_CFOP = CFOP.IDCountCfop where EM.IDestoque = " & TBEstoque!IDEstoque & " and EM.Operacao <> 'SAIDA_RETALHO' and (CPL.Remessa = 0 or CPL.Remessa IS NULL) and (CFOP.MaoObra = 'False' or CFOP.MaoObra IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    qtdeliberada = IIf(IsNull(TBFI!qtdeliberada), 0, TBFI!qtdeliberada)
                    qtdeliberadaPC = IIf(IsNull(TBFI!qtdeliberadaPC), 0, TBFI!qtdeliberadaPC)
                    qtdeliberar = IIf(IsNull(TBFI!qtdeliberar), 0, TBFI!qtdeliberar)
                    qtdeliberarPC = IIf(IsNull(TBFI!qtdeliberarPC), 0, TBFI!qtdeliberarPC)
                    QtdeEstoque = Format(qtdeliberada - qtdeliberar, "###,##0.0000")
                    QtdeEstoquePC = Int(qtdeliberadaPC - qtdeliberarPC)
                End If
                
                'Verifica quantidade de saída do retalho
                If QtdeEstoque > 0 Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select Sum(EM.Saida) as qtdeliberar, Sum(ISNULL(EM.Saida_PC, 0)) as qtdeliberarPC from ((Estoque_movimentacao EM LEFT JOIN Estoque_controle_recebimento ECR ON ECR.Id = EM.IDEstoque_recebimento) LEFT JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido AND CPL.IdLista = ECR.IdLista) LEFT JOIN tbl_NaturezaOperacao CFOP ON CPL.ID_CFOP = CFOP.IDCountCfop where EM.IDestoque = " & TBEstoque!IDEstoque & " and EM.Operacao = 'SAIDA_RETALHO' and (CPL.Remessa = 0 or CPL.Remessa IS NULL) and (CFOP.MaoObra = 'False' or CFOP.MaoObra IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        qtdeliberar = IIf(IsNull(TBFI!qtdeliberar), 0, TBFI!qtdeliberar)
                        qtdeliberarPC = IIf(IsNull(TBFI!qtdeliberarPC), 0, TBFI!qtdeliberarPC)
                        QtdeEstoque = Format(QtdeEstoque - qtdeliberar, "###,##0.0000")
                        QtdeEstoquePC = Int(QtdeEstoquePC - qtdeliberarPC)
                    End If
                    If QtdeEstoque < 0 Then
                        QtdeEstoque = 0
                        QtdeEstoquePC = 0
                    End If
                End If
                TBFI.Close
            Else
                QtdeEstoque = 0
                QtdeEstoquePC = 0
            End If
            TBEstoque!estoque_real = QtdeEstoque
            TBEstoque!estoque_real_PC = QtdeEstoquePC
            TBEstoque!estoque_venda = QtdeEstoque
            TBEstoque!Valor_total = Format(TBEstoque!valor_unitario * QtdeEstoque, "###,##0.00")
            TBEstoque.Update
        Else
            USMsgBox ("Não foi possível atualizar, pois este produto não está cadastrado."), vbExclamation, "CAPRIND v5.0"
            TBEstoque.Close
            TBAbrir.Close
            Exit Sub
        End If
        TBAbrir.Close
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Movimentação"
        Evento = "Atualizar quantidade em estoque do RE"
        ID_documento = Lista.SelectedItem
        Documento = "Cód. interno: " & TBEstoque!Desenho & " - Nº lote: " & TBEstoque!LOTE & " - Nº corrida: " & TBEstoque!Corrida & " - Nº certificado: " & TBEstoque!Certificado & " - Local armaz.: " & TBEstoque!local_armaz
        Documento1 = ""
        ProcGravaEvento
        '==================================
        ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    End If
    TBEstoque.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_empenho_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If Lista.SelectedItem.SubItems(3) = "" Then Exit Sub
Desenho = Lista.SelectedItem.ListSubItems(3)
PCP_Ordem = False
frmEstoque_Empenho.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir_sucata()
On Error GoTo tratar_erro
Dim ID_Antigo As Integer 'OK

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtlocalização.Text = "" Then Exit Sub

Qtde = 0
Qtd = 0
IDlista = 0
Set TBExecucao = CreateObject("adodb.recordset")
TBExecucao.Open "select * from Estoque_Controle where IdEstoque = " & txtlocalização & " and idLote_sucata <> 0 and idLote_sucata IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBExecucao.EOF = False Then
    If TBExecucao!status = "ENTRADA_SUCATA" Then
        MsgTexto = "sucata"
        MsgTexto1 = "Sucata"
    Else
        MsgTexto = "retalho"
        MsgTexto1 = "Retalho"
    End If
    If USMsgBox("Deseja realmente excluir este RE de " & MsgTexto & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        
        'Verifica se existe movimentação de saida e bloqueia exclusão
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select Saida from estoque_movimentacao where idestoque = " & TBExecucao!IDEstoque & " and Saida <> 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = False Then
            USMsgBox ("Não é permitido excluir este RE, pois o mesmo já foi movimentado."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
                
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "select * from estoque_controle where IdEstoque = " & TBExecucao!idLote_sucata & " and idestoque <> " & TBExecucao!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "select * from estoque_controle where idEstoque = " & TBExecucao!idLote_sucata, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBExecucao!status = "ENTRADA_SUCATA" Then
                    If TBAbrir!Desenho <> TBExecucao!Desenho_sucata Then
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "select * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
                        TBGravar.AddNew
                        Set TBAcessos = CreateObject("adodb.recordset")
                        TBAcessos.Open "Select * from estoque_movimentacao where idestoque = " & TBAbrir!IDEstoque & " and desenho = '" & TBAbrir!Desenho & "' and Operacao = 'ENTRADA_SUCATA'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAcessos.EOF = False Then
                            Do While TBAcessos.EOF = False
                                Qtde = Qtde + TBAcessos!Entrada
                                TBAcessos!IDEstoque = TBGravar!IDEstoque
                                TBAcessos.Update
                                TBAcessos.MoveNext
                            Loop
                        End If
                        TBAcessos.Close
                        
                        Set TBItem = CreateObject("adodb.recordset")
                        TBItem.Open "select * from projproduto where desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBItem.EOF = False Then
                            TBGravar!Classe = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
                            TBGravar!valor_unitario = IIf(IsNull(TBItem!PCusto), "", TBItem!PCusto)
                            ValorTotal = IIf(IsNull(TBItem!PCusto), "", TBItem!PCusto)
                            TBGravar!Valor_total = Format(Qtde * ValorTotal, "###,##0.00")
                        End If
                        TBItem.Close
                        TBGravar!estoque_real = Format(Qtde, "###,##0.0000")
                        TBGravar!Qtde = TBGravar!estoque_real
                        TBGravar!estoque_venda = TBGravar!estoque_real
                        TBGravar!Desenho = TBAbrir!Desenho
                        TBGravar!LOTE = TBAbrir!LOTE
                        TBGravar!Descricao = TBAbrir!Descricao
                        TBGravar!Un = TBAbrir!Un
                        TBGravar!idLote_sucata = TBAbrir!IDEstoque
                        TBGravar!Desenho_sucata = TBAbrir!Desenho_sucata
                        TBGravar!Data = Format(Date, "dd/mm/yy")
                        TBGravar!Responsavel = pubUsuario
                        TBGravar!Fornecedor = IIf(IsNull(TBAbrir!Fornecedor), "", TBAbrir!Fornecedor)
                        TBGravar!Certificado = IIf(IsNull(TBAbrir!Certificado), "", TBAbrir!Certificado)
                        TBGravar!local_armaz = IIf(IsNull(TBAbrir!local_armaz), "", TBAbrir!local_armaz)
                        TBGravar!status = IIf(IsNull(TBAbrir!status), "", TBAbrir!status)
                        TBGravar!Corrida = IIf(IsNull(TBAbrir!Corrida), "", TBAbrir!Corrida)
                        If TBAbrir!Consignacao = True Then TBGravar!Consignacao = True
                        TBGravar.Update
                        TBGravar.Close
                        TBAbrir!Desenho = TBAbrir!Desenho_sucata
                        TBAbrir!idLote_sucata = Null
                        TBAbrir!Desenho_sucata = Null
                    End If
                    
                    Qtd = TBAbrir!estoque_real + TBExecucao!estoque_real
                    TBAbrir!estoque_real = Format(Qtd - Qtde, "###,##0.0000")
                Else
                    'Verifica qtde. de saída para o retalho
                    Set TBAcessos = CreateObject("adodb.recordset")
                    TBAcessos.Open "Select Saida from estoque_movimentacao where idestoque = " & TBAbrir!IDEstoque & " and Operacao = 'SAIDA_RETALHO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAcessos.EOF = False Then
                        TBAbrir!estoque_real = TBAbrir!estoque_real + TBAcessos!Saida
                    End If
                    TBAcessos.Close
                End If
                'TBAbrir!Qtde = TBAbrir!estoque_real
                TBAbrir!estoque_venda = TBAbrir!estoque_real
                TBAbrir!Valor_total = Format(TBAbrir!valor_unitario * TBAbrir!estoque_real, "###,##0.00")
                TBAbrir.Update
                
                IDlista = TBAbrir!IDEstoque
            End If
            TBAbrir.Close
            
'            Qtde = 0
'            Set TBAbrir = CreateObject("adodb.recordset")
'            TBAbrir.Open "Select * from estoque_movimentacao where idestoque = " & IDlista & " order by Data desc, IDoperacao desc, Conexao, adOpenKeyset, adLockOptimistic"
'            If TBAbrir.EOF = False Then
'                If TBAbrir!Operacao = "ENTRADA_SUCATA" Then
'                    Qtde = TBAbrir!Entrada
'                    Set TBGravar = CreateObject("adodb.recordset")
'                    TBGravar.Open "select * from estoque_movimentacao where idestoque = " & IDlista & " and operacao = 'ENTRADA_SUCATA'", Conexao, adOpenKeyset, adLockOptimistic
'                    If TBGravar.EOF = False Then
'                        TBGravar!Entrada = Format(TBGravar!Entrada + Qtde, "###,##0.00")
'                        TBGravar.Update
'                    End If
'                    TBGravar.Close
'                End If
'            End If
'            TBAbrir.Close
            
            If TBExecucao!status = "ENTRADA_SUCATA" Then TextoFiltro = "(Operacao = 'ENTRADA_SUCATA' or Operacao = 'SAIDA_SUCATA')" Else TextoFiltro = "(Operacao = 'ENTRADA_RETALHO' or Operacao = 'SAIDA_RETALHO')"
            Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & txtlocalização & " and Data = '" & TBExecucao!Data & "' and " & TextoFiltro
            Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & IDlista & " and Data = '" & TBExecucao!Data & "' and " & TextoFiltro
        Else
            Set TBAcessos = CreateObject("adodb.recordset")
            TBAcessos.Open "Select * from estoque_controle where IdEstoque = " & TBExecucao!idLote_sucata, Conexao, adOpenKeyset, adLockOptimistic
            If TBAcessos.EOF = False Then
                Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & txtlocalização & " and desenho = '" & Lista.SelectedItem.SubItems(3) & "' and lote = '" & Lista.SelectedItem.SubItems(2) & "' and (Operacao = 'ENTRADA_SUCATA' or Operacao = 'SAIDA_SUCATA')"
                Conexao.Execute "DELETE from estoque_movimentacao where IdEstoque = " & TBExecucao!idLote_sucata & " and documento = '" & Lista.SelectedItem.SubItems(3) & "' and lote = '" & Lista.SelectedItem.SubItems(2) & "' and (Operacao = 'ENTRADA_SUCATA' or Operacao = 'SAIDA_SUCATA')"
                
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "select * from projproduto where desenho = '" & TBExecucao!Desenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    TBExecucao!Desenho = TBExecucao!Desenho_sucata
                    TBExecucao!Descricao = TBItem!Descricao
                    TBExecucao!Un = TBItem!Unidade
                    TBExecucao!Classe = TBItem!Classe
                End If
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "select Compras_pedido_lista.* from compras_pedido_lista inner join compras_pedido on compras_pedido_lista.idpedido = compras_pedido.idpedido where compras_pedido.pedido = '" & TBExecucao!LOTE & "' and compras_pedido_lista.desenho = '" & TBExecucao!Desenho_sucata & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then
                    ValorTotal = IIf(IsNull(TBPedido!preco_unitario), "0", TBPedido!preco_unitario)
                End If
                TBPedido.Close
                quantestoque = TBExecucao!estoque_real
                TBExecucao!Valor_total = Format(quantestoque * ValorTotal, "###,##0.00")
                TBExecucao!Desenho_sucata = ""
                TBExecucao!idLote_sucata = 0
            Else
                TBExecucao!Desenho = TBExecucao!Desenho_sucata
                Conexao.Execute "Update estoque_controle Set idLote_sucata = " & TBExecucao!IDEstoque & " where idLote_sucata = " & TBExecucao!idLote_sucata & " and idestoque <> " & txtlocalização
                TBExecucao!idLote_sucata = Null
            End If
            TBExecucao.Update
            TBAcessos.Close
        End If
        TBFIltro.Close
        
        TBExecucao.Delete
        
        USMsgBox (MsgTexto1 & " excluído com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Movimentação"
        Evento = "Excluir " & MsgTexto
        ID_documento = txtlocalização
        Documento = "Cód. interno: " & Lista.SelectedItem.SubItems(3)
        Documento1 = ""
        ProcGravaEvento
        '==================================
        ProcAtualizalista (1)
    End If
Else
    USMsgBox ("Favor selecionar um RE de sucata/retalho antes de excluir."), vbExclamation, "CAPRIND v5.0"
End If
TBExecucao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro
Dim Entrada As Boolean 'OK

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If

Permitido = False
TextoFiltro = ""
With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Permitido = True
            If TextoFiltro = "" Then TextoFiltro = "Idoperacao <> " & .ListItems.Item(InitFor) Else TextoFiltro = TextoFiltro & " and Idoperacao <> " & .ListItems.Item(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) movimentação(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    qtdeliberada = 0
    qtdeliberar = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(Entrada) as qtdeliberada, Sum(Saida) as qtdeliberar from Estoque_movimentacao where IDestoque = " & Lista.SelectedItem & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
        qtdeliberar = IIf(IsNull(TBAbrir!qtdeliberar), 0, TBAbrir!qtdeliberar)
        Valor1 = qtdeliberada - qtdeliberar
        
        Permitido = True
        If Lista.SelectedItem.ListSubItems(13) < 0 Then
            If Valor1 < Lista.SelectedItem.ListSubItems(13) Then Permitido = False
        Else
            If Valor1 < 0 Then Permitido = False
        End If
        If Permitido = False Then
            USMsgBox ("Não é permitido excluir essa(s) movimentação(ões), pois o saldo ficará negativo."), vbExclamation, "CAPRIND v5.0"
            TBAbrir.Close
            Exit Sub
        End If
    End If
    TBAbrir.Close
End If

Permitido = False
With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) movimentação(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            quantidade = 0
            Qtde = 0
                    
            Status_movimentacao = .ListItems(InitFor).SubItems(2)
            'Verif. número do documento
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Estoque_movimentacao where idestoque = " & Lista.SelectedItem & " and (operacao = 'SAIDA_REQUISICAO_PARCIAL' or operacao = 'SAIDA_REQUISICAO')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Documento_Ordem = IIf(IsNull(TBAbrir!Documento), "", TBAbrir!Documento)
            End If
            TBAbrir.Close
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from estoque_movimentacao where idoperacao = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                IDEstoque = TBAbrir!IDEstoque
                
                valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
                
                'Define valor na variável
                If TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Or TBAbrir!Operacao = "ENTRADA_INVENTÁRIO" Or TBAbrir!Operacao = "ENTRADA_DEVOLUÇÃO" Then
                    QuantSolicitado = IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada)
                    
                    'Exclui o empenho no produto em estoque para o pedido
                    Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBAbrir!IDEstoque
                    
                    'Atualiza quantidade de entrada no empenho da ordem
                    If TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Then
                        qtdeliberada = QuantSolicitado
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select PP.Qtde_entrada from (producao_pedidos PP INNER JOIN vendas_carteira VC ON PP.IDcarteira = VC.Codigo) INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Ordem = " & TBAbrir!LOTE & " and P.Desenho = '" & TBAbrir!Desenho & "' and ISNULL(Qtde_entrada , 0) > 0 order by VC.Prazofinal desc", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            Do While TBFI.EOF = False
                                If qtdeliberada >= 0 Then
                                    If qtdeliberada >= TBFI!Qtde_entrada Then
                                        TBFI!Qtde_entrada = 0
                                        qtdeliberada = 0
                                    Else
                                        qtdeliberada = TBFI!Qtde_entrada - qtdeliberada
                                        TBFI!Qtde_entrada = qtdeliberada
                                    End If
                                    TBFI.Update
                                End If
                                TBFI.MoveNext
                            Loop
                        End If
                        TBFI.Close
                    End If
                Else
                    QuantSolicitado = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                End If
                
                If IsNumeric(.ListItems.Item(InitFor).SubItems(8)) = True And (Status_movimentacao = "SAIDA_NOTA" Or Status_movimentacao = "SAIDA_NOTA_PARCIAL") Then
                    'Atualiza qtde. expedida
                    Qtd = .ListItems.Item(InitFor).SubItems(6)
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select PP.Qtdeexpedida, PP.Dataexpedicao, NFPP.* from vendas_carteira PP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.Codigo and NFPP.Codinterno = PP.Desenho where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " order by PP.PrazoFinal", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = False Then
                        Do While TBGravar.EOF = False
                            If Qtd >= TBGravar!qtdeexpedida Then qt = TBGravar!qtdeexpedida Else qt = Qtd
                            TBGravar!qtdeexpedida = TBGravar!qtdeexpedida - qt
                            Qtd = Qtd - qt
                            
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select Data from Estoque_movimentacao where Idoperacao <> " & .ListItems.Item(InitFor) & " and ID_prod_NF = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF) & " and (Operacao = 'SAIDA_NOTA' or Operacao = 'SAIDA_NOTA_PARCIAL') order by Data desc", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFI.EOF = False Then
                                TBGravar!dataexpedicao = TBFI!Data
                            Else
                                TBGravar!dataexpedicao = Null
                            End If
                            TBFI.Close
                            TBGravar.Update
                            
                            'Desvincula pedido da ordem para estoque
                            If IsNumeric(TBAbrir!LOTE) = True Then
                                Set TBAliquota = CreateObject("adodb.recordset")
                                TBAliquota.Open "Select * from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAliquota.EOF = False Then
                                    TBAliquota!Qtde_empenho = TBAliquota!Qtde_empenho - qt
                                    TBAliquota!Qtde_entrada = TBAliquota!Qtde_empenho
                                    TBAliquota.Update
                                    
                                    If TBAliquota!Qtde_empenho <= 0 Then Conexao.Execute "DELETE from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'"
                                End If
                                TBAliquota.Close
                            End If
                            
                            Do While qt > 0
                                'Atualiza qtde. de saída no empenho
                                Set TBAliquota = CreateObject("adodb.recordset")
                                TBAliquota.Open "Select EE.Qtde_saida from Estoque_Controle_Empenho_Vendas EE INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON EE.ID_carteira = NFPP.ID_carteira where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " and EE.ID_estoque = " & TBAbrir!IDEstoque & " and EE.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                                If TBAliquota.EOF = False Then
                                    If TBAliquota!Qtde_saida >= qt Then
                                        TBAliquota!Qtde_saida = TBAliquota!Qtde_saida - qt
                                        qt = 0
                                    Else
                                        qt = qt - TBAliquota!Qtde_saida
                                        TBAliquota!Qtde_saida = 0
                                    End If
                                    TBAliquota.Update
                                Else
                                    GoTo Prosseguir
                                End If
                                TBAliquota.Close
                            Loop
Prosseguir:
                            If Qtd <= 0 Then GoTo Prosseguir1
                            TBGravar.MoveNext
                        Loop
                    End If
                End If

Prosseguir1:
                If Status_movimentacao = "SAIDA_REQUISICAO" Or Status_movimentacao = "SAIDA_REQUISICAO_PARCIAL" Then
                    Set TBMateriaprima = CreateObject("adodb.recordset")
                    TBMateriaprima.Open "Select * from Requisicao_materiais where requisicao = '" & TBAbrir!Documento & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBMateriaprima.EOF = False Then
                        
                        NovoValor = Replace(QuantSolicitado, ",", ".")
                        
                        Set TBMaterial = CreateObject("adodb.recordset")
                        TBMaterial.Open "Select RML.*, P.ID_PC from Requisicao_materiais_lista RML INNER JOIN Projproduto P ON P.Desenho = RML.Desenho where RML.idrequisicao = " & TBMateriaprima!ID & " and RML.desenho = '" & Lista.SelectedItem.SubItems(3) & "' and RML.quant_saida >= " & NovoValor & " and (RML.Status = 'RETIRADO' or RML.Status = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBMaterial.EOF = False Then
                            Qtde = IIf(IsNull(TBMaterial!quant_saida), 0, TBMaterial!quant_saida)
                            Qtde = Qtde - QuantSolicitado
                            TBMaterial!quant_saida = Format(Qtde, "###,##0.0000")
                            If Qtde > 0 Then TBMaterial!status = "PARCIAL" Else TBMaterial!status = "REQUISIT."
                            TBMaterial.Update
                            
                            If IsNull(TBMaterial!Ordem) = False And TBMaterial!Ordem <> 0 And IsNull(TBMaterial!ID_PC) = False And TBMaterial!ID_PC <> 0 Then
                                Set TBOrdem = CreateObject("adodb.recordset")
                                TBOrdem.Open "Select * from Producao_outras_despesas where Ordem = " & TBMaterial!Ordem & " and ID_PC = " & TBMaterial!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                                If TBOrdem.EOF = False Then
                                    If TBOrdem!valor - (Format(valor * QuantSolicitado, "###,##0.00")) <= 0 Then
                                        TBOrdem.Delete
                                    Else
                                        TBOrdem!valor = TBOrdem!valor - (Format(valor * QuantSolicitado, "###,##0.00"))
                                        TBOrdem.Update
                                    End If
                                End If
                                Valor1 = 0
                                Set TBOrdem = CreateObject("adodb.recordset")
                                TBOrdem.Open "Select Sum(Valor) as Valor1 from Producao_outras_despesas where Ordem = " & TBMaterial!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                                If TBOrdem.EOF = False Then
                                    Valor1 = IIf(IsNull(TBOrdem!Valor1), 0, TBOrdem!Valor1)
                                End If
                                TBOrdem.Close
                                NovoValor = Replace(Valor1, ",", ".")
                                Conexao.Execute "Update Producao Set CTOutras = " & NovoValor & " where Ordem = " & TBMaterial!Ordem
                                
                            End If
                        End If
                        TBMaterial.Close
                        
                        ProcAtualizaStatus_RM
                    End If
                End If
                
                If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Or Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                    If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Then
                        'Achando a movimentação de saida
                        Set TBNivel14 = CreateObject("adodb.recordset")
                        TBNivel14.Open "Select IdEstoque, idoperacao from Estoque_movimentacao where idoperacao = " & TBAbrir!IdTrocaLocal, Conexao, adOpenKeyset, adLockOptimistic
                        If TBNivel14.EOF = False Then
                            'achando o RE de saida
                            Set TBMateriaprima = CreateObject("adodb.recordset")
                            TBMateriaprima.Open "Select estoque_real_PC, estoque_real, IDestoque from Estoque_controle where IdEstoque = " & TBNivel14!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                            If TBMateriaprima.EOF = False Then
                                procVoltarEmpenhoLocal TBAbrir!IDEstoque, TBMateriaprima!IDEstoque
                                TBMateriaprima!estoque_real = TBMateriaprima!estoque_real + TBAbrir!Entrada
                                TBMateriaprima!estoque_real_PC = TBMateriaprima!estoque_real_PC + TBAbrir!Entrada_PC
                                TBMateriaprima.Update
                                TBNivel14.Delete 'Exclui movimentação de saida
                            End If
                            TBMateriaprima.Close
                        End If
                        TBNivel14.Close
                    Else
                        'Achando a movimentação de entrada
                        Set TBNivel14 = CreateObject("adodb.recordset")
                        TBNivel14.Open "Select IdEstoque, idoperacao from Estoque_movimentacao where IdTrocaLocal = " & TBAbrir!IDoperacao, Conexao, adOpenKeyset, adLockOptimistic
                        If TBNivel14.EOF = False Then
                            procVoltarEmpenhoLocal TBNivel14!IDEstoque, TBAbrir!IDEstoque
                            Conexao.Execute "DELETE FROM Estoque_controle WHERE IdEstoque = " & TBNivel14!IDEstoque 'Exclui estoque controle entrada
                            TBNivel14.Delete 'Exclui movimentação entrada
                        End If
                        TBNivel14.Close
                    End If
                End If
            End If
            TBAbrir.Close
            Conexao.Execute "DELETE from estoque_movimentacao where idoperacao = " & .ListItems.Item(InitFor)
            
            'Corrige retirada na tabela producaomaterial
            quantidade = 0
            QuantidadePC = 0
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from Estoque_controle where IDEstoque = " & Lista.SelectedItem & " and consignacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                TextoFiltro = "IDEstoque = " & Lista.SelectedItem
            Else
                TextoFiltro = "Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "'"
            End If
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Saida) as Quantidade, Sum(ISNULL(Saida_PC, 0)) as QuantidadePC from estoque_movimentacao where " & TextoFiltro & " and documento = '" & .ListItems.Item(InitFor).ListSubItems(8) & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                quantidade = IIf(IsNull(TBAbrir!quantidade), 0, TBAbrir!quantidade)
                QuantidadePC = IIf(IsNull(TBAbrir!QuantidadePC), 0, TBAbrir!QuantidadePC)
            End If
            TBAbrir.Close
            
            If IsNumeric(.ListItems.Item(InitFor).SubItems(8)) = True Then
                Set TBproducao = CreateObject("adodb.recordset")
                TBproducao.Open "Select * from producaomaterial where codigo = '" & Lista.SelectedItem.ListSubItems(3) & "' and Ordem = " & .ListItems.Item(InitFor).SubItems(8), Conexao, adOpenKeyset, adLockOptimistic
                If TBproducao.EOF = False Then
                    If quantidade = 0 Then
                        TBproducao!Saida = "NÃO"
                    ElseIf quantidade >= TBproducao!Requisitado Or QuantidadePC >= TBproducao!Total_pc Then
                        TBproducao!Saida = "SIM"
                    Else
                        TBproducao!Saida = "PARCIAL"
                    End If
                    
                    TBproducao!Valor_saida_estoque = Format(IIf(IsNull(TBproducao!Valor_saida_estoque), 0, TBproducao!Valor_saida_estoque) - (valor * QuantSolicitado), "###,##0.00")
                    TBproducao.Update
                    
                    'Atualiza qtde. de saída do empenho da ordem
                    QuantEmpenho = 0
                    QuantEmpenhoPC = 0
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select Sum(Saida) as QuantEmpenho, Sum(ISNULL(Saida_PC, 0)) as QuantEmpenhoPC from estoque_movimentacao where IDestoque = " & Lista.SelectedItem & " and oe = '" & TBproducao!Ordem & "' and desenho = '" & TBproducao!CODIGO & "' and documento = '" & TBproducao!Ordem & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, Format(TBAbrir!QuantEmpenho, "###,##0.0000"))
                        QuantEmpenhoPC = IIf(IsNull(TBAbrir!QuantEmpenhoPC), 0, Format(TBAbrir!QuantEmpenhoPC, "###,##0.0000"))
                    End If
                    NovoValor = Replace(QuantEmpenho, ",", ".")
                    NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
                    Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & Lista.SelectedItem & " and Ordem = " & TBproducao!Ordem & " and Codinterno = '" & TBproducao!CODIGO & "'"
                    Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & Lista.SelectedItem & " and Ordem = " & TBproducao!Ordem & " and Codinterno = '" & TBproducao!CODIGO & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
                End If
                TBproducao.Close
            End If
            
            Permitido1 = True
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from projproduto where Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = True Then
                Permitido1 = False
            End If
            TBProduto.Close
            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & .ListItems.Item(InitFor)
            
            'Corrige estoque real
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle where idestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                
                '==================================
                Modulo = "Estoque/Movimentação"
                Evento = "Excluir movimentação"
                ID_documento = .ListItems.Item(InitFor)
                Documento = "Cód. interno: " & TBEstoque!Desenho & " - Nº lote: " & TBEstoque!LOTE & " - Nº corrida: " & TBEstoque!Corrida & " - Nº certificado: " & TBEstoque!Certificado & " - Local armaz.: " & TBEstoque!local_armaz
                Documento1 = "Operação: " & .ListItems.Item(InitFor).SubItems(2) & " - Data: " & .ListItems.Item(InitFor).SubItems(3) & " - Entrada: " & .ListItems.Item(InitFor).SubItems(4) & " - Saída: " & .ListItems.Item(InitFor).SubItems(6) & " - Documento: " & .ListItems.Item(InitFor).SubItems(8)
                ProcGravaEvento
                '==================================
                
                If Permitido1 = True Then
                    Qtde = Valor1
                    TBEstoque!estoque_real = Format(Qtde, "###,##0.0000")
                    TBEstoque!Qtde = Format(Qtde, "###,##0.0000")
                    TBEstoque!estoque_real_PC = Format(IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - IIf(.ListItems.Item(InitFor).SubItems(5) = "", 0, .ListItems.Item(InitFor).SubItems(5)) + IIf(.ListItems.Item(InitFor).SubItems(7) = "", 0, .ListItems.Item(InitFor).SubItems(7)), "###,##0.0000")
                Else
                    TBEstoque!estoque_real = 0
                    TBEstoque!estoque_real_PC = 0
                    Qtde = 0
                End If
                                        
                'Atualiza valor do material no estoque
                'Estoque_controle
                TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Qtde, "###,##0.00")
                
                TBEstoque.Update
                Set TBMaterial = CreateObject("adodb.recordset")
                TBMaterial.Open "Select * from Estoque_movimentacao where IDEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBMaterial.EOF = True Then TBEstoque.Delete
                TBMaterial.Close
            End If
            TBEstoque.Close
            
            If Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Then
                Conexao.Execute "Update estoque_controle Set Status = 'ENTRADA_ORDEM_PARCIAL' where lote = '" & Lista.SelectedItem.SubItems(3) & "' and status = 'ENTRADA_ORDEM' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                Conexao.Execute "Update estoque_movimentacao Set Operacao = 'ENTRADA_ORDEM_PARCIAL' where lote = '" & Lista.SelectedItem.SubItems(3) & "' and Operacao = 'ENTRADA_ORDEM'"
            End If
            
            If IsNumeric(.ListItems.Item(InitFor).SubItems(8)) = True And (Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Or Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL") Then
                'Atualiza qtde. produzida
                Qtde = 0
                qtdeliberada = 0
                Qtd = 0
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from producao where Ordem = " & .ListItems.Item(InitFor).SubItems(8) & " and Controlado_estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Then Qtd = .ListItems.Item(InitFor).SubItems(4)
                    If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then Qtd = .ListItems.Item(InitFor).SubItems(6)
                    Set TBExecucao = CreateObject("adodb.recordset")
                    TBExecucao.Open "select * from estoque_controle where Lote = '" & Lista.SelectedItem.SubItems(3) & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao.EOF = False Then
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select Sum(Entrada) as Qtde, Sum(Saida) as qtdeliberada from estoque_movimentacao where idestoque = " & TBExecucao!IDEstoque & " and documento = '" & .ListItems.Item(InitFor).SubItems(8) & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            Qtde = IIf(IsNull(TBCorretiva!Qtde), 0, TBCorretiva!Qtde) + IIf(IsNull(TBCorretiva!qtdeliberada), 0, TBCorretiva!qtdeliberada)
                        End If
                        TBCorretiva.Close
                    End If
                    TBExecucao.Close
                    If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Then
                        If Qtde < Qtd Then ProcOrdem
                    End If
                    If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
                        If Qtde = 0 Then ProcOrdem
                    End If
                End If
                TBCiclo.Close
                
                'Custo material
                If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), .ListItems.Item(InitFor).SubItems(8)
            End If
        End If
     Next InitFor
End With

USMsgBox ("Movimentação(ões) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
    Lista.SelectedItem = Lista.ListItems(CodigoLista)
    ProcCarregaDadosLote
    Lista.SetFocus
Else
2:
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 2
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcLocalArmazenamento()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "alterar"
If txtlocalização = "" Then
    NomeCampo = "o lote na lista"
    ProcVerificaAcao
    Exit Sub
End If
If Lista.SelectedItem.SubItems(8) = "" Then
    NomeCampo = "o local de armazenamento"
    ProcVerificaAcao
    Exit Sub
End If
frmEstoque_item_localarmaz.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSucata()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista.ListItems.Count = "0" Or txtlocalização.Text = "" Then
    USMsgBox ("Infome o RE antes de gerar sucata/retrabalho."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select IDestoque from estoque_controle where idestoque = " & txtlocalização & " and idlote_sucata <> 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é permitido gerar sucata/retrabalho, pois o mesmo já é uma sucata."), vbExclamation, "CAPRIND v5.0"
    TBAbrir.Close
    Exit Sub
End If
TBAbrir.Close
RE = txtlocalização.Text
IDempresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

frmEstoque_sucata.Show 1

'With frmestoque_item
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        ProcCarregaDadosLote
        Lista.SetFocus
    Else
        ProcFiltrar
    End If
'End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If txtlocalização = "" Then
    USMsgBox ("Informe o RE antes de alterar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmEstoque_item_bloq.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEstruturadoitem()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If Lista.SelectedItem.SubItems(3) = "" Then
    USMsgBox ("Informe o código interno antes de abrir a estrutura."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & Lista.SelectedItem.SubItems(3) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    With frmproj_conjunto
        .Show
        .ProcLimpaCampos
        .Txt_cod_produto = TBItem!Codproduto
        .txtdesenhoproduto.Text = TBItem!Desenho
        .txtDescricaoProduto.Text = TBItem!Descricao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from item_aplicacoes where codproduto = " & TBItem!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then .txtRef.Text = TBAbrir("n_referencia")
        TBAbrir.Close
        .ProcAtualizalista (1)
    End With
Else
    USMsgBox ("Não foi encontrado nenhum registro para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_empenhoRE_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If Lista.SelectedItem.SubItems(3) = "" Then Exit Sub
Desenho = ""
IDlista = Lista.SelectedItem
PCP_Ordem = False
frmEstoque_Empenho.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_salvar_ref_serie_Click()
On Error GoTo tratar_erro

Acao = "salvar"
If txtlocalização = "" Then
    NomeCampo = "o RE"
    ProcVerificaAcao
    txtlocalização.SetFocus
    Exit Sub
End If
If USMsgBox("Deseja realmente alterar o código de referência e número de série deste RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    'Verifica se o código de referencia está vinculado a outro produto
    'If Txt_cod_ref <> "" Then If FunVerifiCodRefUtilizado(Lista.SelectedItem.ListSubItems(3), Txt_cod_ref) = True Then Exit Sub
    
    Conexao.Execute "Update estoque_controle Set Ref = '" & Txt_cod_ref & "', Numero_serie = '" & Txt_n_serie & "' where IDestoque = " & txtlocalização
    Conexao.Execute "Update Estoque_fisico Set Cod_ref = '" & Txt_cod_ref & "', Numero_serie = '" & Txt_n_serie & "' where IDestoque = " & txtlocalização
    Conexao.Execute "Update EF set EF.Cod_ref = '" & Txt_cod_ref & "', EF.Numero_serie = '" & Txt_n_serie & "' from Estoque_fisico EF INNER JOIN Estoque_movimentacao EM ON EM.ID_inventario = EF.ID where EM.IDestoque = " & txtlocalização & " and EM.ID_inventario IS NOT NULL and EM.ID_inventario <> 0"
    
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Movimentação"
    Evento = "Alterar código de referência e número de série"
    ID_documento = txtlocalização
    Documento = "Cód. interno: " & Lista.SelectedItem.SubItems(3)
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Lista_Movimentacao.ListItems.Clear
    ProcAtualizalista (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
    If Lista.ListItems.Count <> 0 And CodigoLista <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Movimentacao.AbsolutePage <> 2 Then
    If TBLISTA_Estoque_Movimentacao.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Estoque_Movimentacao.PageCount - 1)
    Else
        TBLISTA_Estoque_Movimentacao.AbsolutePage = TBLISTA_Estoque_Movimentacao.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Estoque_Movimentacao.AbsolutePage)
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
    TBLISTA_Estoque_Movimentacao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Estoque_Movimentacao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Movimentacao.AbsolutePage = 1
ProcExibePagina (TBLISTA_Estoque_Movimentacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Estoque_Movimentacao.AbsolutePage <> -3 Then
    If TBLISTA_Estoque_Movimentacao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Estoque_Movimentacao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Estoque_Movimentacao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Estoque_Movimentacao.AbsolutePage = TBLISTA_Estoque_Movimentacao.PageCount
ProcExibePagina (TBLISTA_Estoque_Movimentacao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcEstruturadoitem
    Case vbKeyF8: ProcSucata
    Case vbKeyF9: ProcExcluir_sucata
    Case vbKeyF10: ProcLocalArmazenamento
    Case vbKeyF11: ProcAlterar_valor
    Case vbKeyF12: ProcCC
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCC()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Lista_Movimentacao.ListItems.Count = 0 Then
    USMsgBox ("Informe a movimentação antes de visualizar o(s) centro(s) de custo."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Status_movimentacao = Lista_Movimentacao.SelectedItem.ListSubItems(2)
If Status_movimentacao = "SAIDA_REQUISICAO" Or Status_movimentacao = "SAIDA_REQUISICAO_PARCIAL" Or Status_movimentacao = "ENTRADA_INVENTÁRIO" Or Status_movimentacao = "SAIDA_INVENTÁRIO" Or Status_movimentacao = "ENTRADA_DEVOLUÇÃO" Then
    Estoque_recebimento = False
Else
    Estoque_recebimento = True
End If
frmEstoque_item_lista_CC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 17, True
Formulario = "Estoque/Movimentação"
txtvencimento.Value = Date

Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, False

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "Produtos/Serviços", "T", True
If Permitido = False Then cmbfiltrarpor = "Código interno"

ProcRemoveObjetosResize Me
'ProcCorrigeSaldoRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Estoque/Movimentação"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmestoque_item_imprimir.Show 1

'NomeRel = "Estoque_saldo_resumido2.rpt"
'Debug.print FormulaRelatorio
'ProcImprimirRel FormulaRelatorio, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procAtualiza()
On Error GoTo tratar_erro
LiberarAlteracao = False

frmSenha.Show 1

If LiberarAlteracao = True Then
frmestoque_item_atualizar.Show 1
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizacao()
On Error GoTo tratar_erro
Dim EntradaPC As Double, Saida As Double, SaidaPC As Double, Total As Double, TotalPC As Double 'OK

Entrada = 0
EntradaPC = 0
Saida = 0
SaidaPC = 0
Total = 0
valor = 0
Valor1 = 0
Valor2 = 0

If USMsgBox("Deseja realmente executar essa(s) atualização(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmestoque_item_atualizar
        If .chkSaldoRE.Value = 1 Then
        Conexao.Execute ("update Estoque_Controle Set Estoque_real = ECSRE.Saldo, Estoque_venda = ECSRE.Saldo, Qtde_fisica = ECSRE.saldo from Estoque_Controle EC inner Join Estoque_controle_Saldo_RE ECSRE on EC.IdEstoque = ECSRE.RE where EC.IdEstoque = ECSRE.RE ")
        USMsgBox "Acerto de saldo de RE executado com sucesso!", vbInformation, "CAPRIND v5.0"
        End If
        
        If .Chk1.Value = 1 Then
            'Atualiza estoque controle de movimentação sem estoque controle
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_controle where IDestoque = " & TBEstoque!IDEstoque & " or Lote = '" & TBEstoque!LOTE & "' and Desenho = '" & TBEstoque!Desenho & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        TBAbrir.AddNew
                        TBAbrir!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                        TBAbrir!LOTE = TBEstoque!LOTE
                        TBAbrir!Desenho = TBEstoque!Desenho
                        TBAbrir!Descricao = TBEstoque!Descricao
                        TBAbrir!estoque_venda = 0
                        TBAbrir!estoque_real = 0
                        TBAbrir!estoque_real_PC = 0
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            TBAbrir!Un = TBFI!Unidade
                            TBAbrir!Classe = TBFI!Classe
                        End If
                        TBFI.Close
                        
                        TBAbrir!Data = Date
                        TBAbrir!Responsavel = pubUsuario
                        
                        Set TBPedido = CreateObject("adodb.recordset")
                        TBPedido.Open "Select * from compras_pedido where Pedido = '" & TBEstoque!LOTE & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBPedido.EOF = False Then
                            
                            TBAbrir!status = "ENTRADA_NOTA_FISCAL"
                            
                            Set TBCompras = CreateObject("adodb.recordset")
                            TBCompras.Open "Select * from Estoque_controle_recebimento where ID = " & TBEstoque!IDEstoque_recebimento, Conexao, adOpenKeyset, adLockOptimistic
                            If TBCompras.EOF = False Then
                                TBAbrir!Certificado = TBCompras!Certificado
                                TBAbrir!Corrida = TBCompras!Corrida
                                TBAbrir!local_armaz = TBCompras!local_armaz
                            Else
                                GoTo 1:
                            End If
                            TBCompras.Close
                        Else
                            If IsNumeric(TBEstoque!LOTE) = True Then
                                Set TBproducao = CreateObject("adodb.recordset")
                                TBproducao.Open "Select * from Producao where Ordem = " & TBEstoque!LOTE, Conexao, adOpenKeyset, adLockOptimistic
                                If TBproducao.EOF = False Then
                                    TBAbrir!status = "ENTRADA_ORDEM"
                                Else
                                    TBAbrir!status = "ENTRADA_INVENTÁRIO"
                                End If
                                TBproducao.Close
                            Else
                                TBAbrir!status = "ENTRADA_INVENTÁRIO"
                            End If
1:
                            TBAbrir!Certificado = 0
                            TBAbrir!Corrida = 0
                            
                            Set TBFIltro = CreateObject("adodb.recordset")
                            TBFIltro.Open "Select * from estoque_controle where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFIltro.EOF = False Then
                                TBAbrir!local_armaz = TBFIltro!local_armaz
                            Else
                                TBAbrir!local_armaz = "N/C"
                            End If
                            TBFIltro.Close
                        End If
                        TBPedido.Close
                        
                        TBAbrir.Update
                        TBEstoque!IDEstoque = TBAbrir!IDEstoque
                        TBEstoque.Update
                        
                    'Else
                        'If TBEstoque!IdEstoque <> TBAbrir!IdEstoque Then
                            'Conexao.Execute "Update Estoque_movimentacao Set idestoque = " & TBAbrir!IdEstoque & " where IDestoque = " & TBEstoque!IdEstoque
                        'End If
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_Controle order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_Controle Where Desenho = '" & TBEstoque!Desenho & "' and Lote = '" & TBEstoque!LOTE & "' and Corrida = '" & TBEstoque!Corrida & "' and Certificado = '" & TBEstoque!Certificado & "' and local_armaz = '" & TBEstoque!local_armaz & "' and Idestoque <> " & TBEstoque!IDEstoque & " and ID_empresa = " & TBEstoque!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            TBEstoque!estoque_venda = Format(TBEstoque!estoque_venda + TBAbrir!estoque_venda, "###,##0.0000000000")
                            TBEstoque!estoque_real = Format(TBEstoque!estoque_real + TBAbrir!estoque_real, "###,##0.0000000000")
                            TBEstoque!estoque_real_PC = Format(TBEstoque!estoque_real_PC + TBAbrir!estoque_real_PC, "###,##0.0000000000")
                            TBEstoque!Qtde = Format(TBEstoque!Qtde + TBAbrir!Qtde, "###,##0.0000000000")
                            TBEstoque.Update
                            
                            Conexao.Execute "Update Estoque_movimentacao Set IdEstoque = " & TBEstoque!IDEstoque & " where Idestoque = " & TBAbrir!IDEstoque
                            Conexao.Execute "DELETE from Estoque_Controle where Idestoque = " & TBAbrir!IDEstoque
                            TBAbrir.MoveNext
                        Loop
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            
            'Deleta movimentação sem estoque_controle
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_controle where Idestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        Conexao.Execute "DELETE from Estoque_movimentacao where Idestoque = " & TBEstoque!IDEstoque
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
            
        If .Chk2.Value = 1 Then
            'Custo material na ordens
            Conexao.Execute "Update producao Set CPR = 0, CTTReal = 0, CTMaterial = 0, CTServico = 0"
            
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from Producao order by Tipo desc, Ordem", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                TBCiclo.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBCiclo.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBCiclo.MoveFirst
                Do While TBCiclo.EOF = False
                    valor = 0
                    Valor1 = 0
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select IDlista from Compras_pedido_lista where Ordem = " & TBCiclo!Ordem & " and Tipo = 'P' and (Status_Item = 'N_RECEBIDO' or Status_Item = 'RECEBIDO' or Status_Item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select CODIGO from producaomaterial where Ordem = " & TBCiclo!Ordem & " order by codigo", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            Do While TBFI.EOF = False
                                Set TBEstoque = CreateObject("adodb.recordset")
                                TBEstoque.Open "Select Sum(EM.VlrTotal) as Valor from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.Documento = '" & TBCiclo!Ordem & "' and EM.Desenho = '" & TBFI!CODIGO & "' and EC.Consignacao = 'False' and (EM.Operacao = 'SAIDA_ORDEM' or EM.Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                                If TBEstoque.EOF = False Then
                                    valor = IIf(IsNull(TBEstoque!valor), 0, TBEstoque!valor)
                                End If
                                TBFI!Valor_saida_estoque = valor
                                Valor1 = Valor1 + valor
                                TBFI.Update
                                
                                TBEstoque.Close
                                TBFI.MoveNext
                            Loop
                        End If
                        TBFI.Close
                        TBCiclo!CTMaterial = Format(Valor1, "###,##0.00")
                        TBCiclo.Update
                    End If
                    TBAbrir.Close
                    TBCiclo.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
        End If
        
        If .Chk3.Value = 1 Then
            'Atualiza valor unitário do iventário que esta zerado
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_movimentacao where VlrUnit = 0 and (Operacao = 'ENTRADA_INVENTÁRIO' or Operacao = 'SAIDA_INVENTÁRIO') order by idEstoque, Data", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_movimentacao where idEstoque = " & TBEstoque!IDEstoque & " and Idoperacao <> " & TBEstoque!IDoperacao & " and VlrUnit <> 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        TBEstoque!VlrUnit = TBAbrir!VlrUnit
                        If TBEstoque!Operacao = "ENTRADA_INVENTÁRIO" Then
                            TBEstoque!vlrTotal = TBAbrir!VlrUnit * TBEstoque!Entrada
                        Else
                            TBEstoque!vlrTotal = TBAbrir!VlrUnit * TBEstoque!Saida
                        End If
                    End If
                    TBAbrir.Close
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBEstoque.Close
            
            'Valor unitário de entrada e local de armazenamento na lista de materias da ordem
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from Estoque_Controle order by idEstoque", Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                TBEstoque.MoveLast
                PBLista.Min = 0
                PBLista.Max = TBEstoque.RecordCount
                PBLista.Value = 1
                Contador = 0
                TBEstoque.MoveFirst
                Do While TBEstoque.EOF = False
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        If TBProduto!Estoque = True Then ControlaEstoque = True Else ControlaEstoque = False
                        TBEstoque!Un = TBProduto!Unidade
                    End If
                    TBProduto.Close
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Estoque_movimentacao where idEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        Do While TBAbrir.EOF = False
                            
                            'Verificar se o produto é remessa e marca como não controla estoque
                            Set TBProduto = CreateObject("adodb.recordset")
                            TBProduto.Open "Select CPL.IDlista from Estoque_controle_recebimento ECR INNER JOIN Compras_pedido_lista CPL ON ECR.IDPedido = CPL.IDPedido and ECR.IdLista = CPL.IdLista and ECR.Desenho = CPL.Desenho where ECR.Id = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento) & " and CPL.remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBProduto.EOF = False Then ControlaEstoque = False
                            TBProduto.Close
                                                        
                            If TBAbrir!Operacao <> "DEVOLUCAO_ALMOXARIFADO C/ PROB." Then
                                Entrada = Entrada + IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada)
                                EntradaPC = EntradaPC + IIf(IsNull(TBAbrir!Entrada_PC), 0, TBAbrir!Entrada_PC)
                            End If
                                
                            Saida = Saida + IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            SaidaPC = SaidaPC + IIf(IsNull(TBAbrir!Saida_PC), 0, TBAbrir!Saida_PC)
                            
                            If TBAbrir!Operacao = "ENTRADA_INVENTÁRIO" Then
                                If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                    'Verif. valor unitário no cadastro do produto
                                    Set TBProduto = CreateObject("adodb.recordset")
                                    TBProduto.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBProduto.EOF = False Then
                                        valor = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
                                    End If
                                    TBProduto.Close
                                Else
                                    valor = TBAbrir!VlrUnit
                                End If
                            ElseIf TBAbrir!Operacao = "ENTRADA_ORDEM" Or TBAbrir!Operacao = "ENTRADA_ORDEM_PARCIAL" Then
                                    'Verif. valor unitário na ordem
                                    Set TBProduto = CreateObject("adodb.recordset")
                                    TBProduto.Open "Select Ordem, Quant, QuantProd, QuantNC, CTTReal, CPR, CTServico, CTMaterial, CTOutras, consignacao from producao where Ordem = " & TBAbrir!LOTE, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBProduto.EOF = False Then
                                                                      'ORDEM           QTDE. PREVISTA                                    QTDE. OK                                                  QT. PROD.(OK+NC)                                                                                                 CUSTO LOTE                                            CUSTO PEÇA                                    CUSTO TERCEIROS                                           CUSTO MATERIAL                                              CUSTO OUTRAS                                            ORDEM CONSIGNADA
                                        valor = FunCalculaValorUnitOrdem(TBProduto!Ordem, IIf(IsNull(TBProduto!Quant), 0, TBProduto!Quant), IIf(IsNull(TBProduto!QuantProd), 0, TBProduto!QuantProd), IIf(IsNull(TBProduto!QuantProd), 0, TBProduto!QuantProd) + IIf(IsNull(TBProduto!QuantNC), 0, TBProduto!QuantNC), IIf(IsNull(TBProduto!CTTReal), 0, TBProduto!CTTReal), IIf(IsNull(TBProduto!CPR), 0, TBProduto!CPR), IIf(IsNull(TBProduto!CTServico), 0, TBProduto!CTServico), IIf(IsNull(TBProduto!CTMaterial), 0, TBProduto!CTMaterial), IIf(IsNull(TBProduto!CTOutras), 0, TBProduto!CTOutras), TBProduto!Consignacao)
                                        OF = TBProduto!Ordem
                                    End If
                                    TBProduto.Close
                                ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Then
                                        If IsNull(TBAbrir!VlrUnit) = True Or TBAbrir!VlrUnit = 0 Then
                                            'Verif. valor unitário no cadastro do produto
                                            Set TBProduto = CreateObject("adodb.recordset")
                                            TBProduto.Open "Select PCusto from projproduto where Desenho = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                            If TBProduto.EOF = False Then
                                                valor = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
                                            End If
                                            TBProduto.Close
                                        End If
                                    ElseIf TBAbrir!Operacao = "ENTRADA_NOTA_FISCAL" Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                            TBFIltro.Open "Select IDlista, ID_empresa from Estoque_controle_recebimento where ID = " & IIf(IsNull(TBAbrir!IDEstoque_recebimento), 0, TBAbrir!IDEstoque_recebimento), Conexao, adOpenKeyset, adLockOptimistic
                                            If TBFIltro.EOF = False Then
                                                
                                                'Verifica dados da NF
                                                Set TBFI = CreateObject("adodb.recordset")
                                                TBFI.Open "Select NF.ID_empresa, NF.Estado, NFP.Int_codigo, NFP.txt_Unid, NFP.Unidade_com, NFP.int_Qtd, NFP.Valor_frete, NFP.Valor_seguro, NFP.Valor_acessorias, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.dbl_ValorUnitario, NFP.txt_Unid, NFP.Unidade_com from (tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NFP.ID_nota = NF.ID) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_prod_NF = NFP.Int_codigo where NFPP.ID_carteira = " & TBFIltro!IDlista & " and NFPP.Codinterno = '" & TBAbrir!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                                                If TBFI.EOF = False Then
                                                    qt = 1
                                                    If TBFI!txt_Unid <> TBFI!Unidade_com And TBFI!Qtde_estoque > 0 Then qt = TBFI!int_Qtd / TBAbrir!Entrada
                                                    
                                                    'Verifica valor do ICMS
                                                    ValorICMS = 0
                                                    Valor1 = 0
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Valor_ICMS, Valor_ICMS_ST, Valor_ICMS_SN from tbl_Detalhes_Nota_CST_ICMS where ID_item = " & TBFI!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If IsNull(TBAliquota!Valor_ICMS) = False And TBAliquota!Valor_ICMS <> 0 Then
                                                            ValorICMS = TBAliquota!Valor_ICMS
                                                        ElseIf IsNull(TBAliquota!Valor_ICMS_ST) = False And TBAliquota!Valor_ICMS_ST <> 0 Then
                                                                ValorICMS = TBAliquota!Valor_ICMS_ST
                                                            ElseIf IsNull(TBAliquota!Valor_ICMS_SN) = False And TBAliquota!Valor_ICMS_SN <> 0 Then
                                                                    ValorICMS = TBAliquota!Valor_ICMS_SN
                                                        End If
                                                    End If
                                                    If ValorICMS <> 0 Then Valor1 = Format(qt * (ValorICMS / TBFI!int_Qtd), "0.0000000000") 'Valor unitário de ICMS
                                                    
                                                    QuantsolicitadoN2 = Format(qt * (IIf(IsNull(TBFI!Valor_desconto), 0, TBFI!Valor_desconto) / TBFI!int_Qtd), "0.0000000000") 'Valor unitário de desconto
                                                    Valor2 = Format(qt * (TBFI!Valor_frete / TBFI!int_Qtd), "0.0000000000")
                                                    ValorPagar = Format(qt * (TBFI!Valor_seguro / TBFI!int_Qtd), "0.0000000000")
                                                    ValorPago = Format(qt * (TBFI!Valor_acessorias / TBFI!int_Qtd), "0.0000000000")
                                                    
                                                    Set TBAliquota = CreateObject("adodb.recordset")
                                                    TBAliquota.Open "Select Simples, Real from Empresa where Codigo = " & TBFIltro!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBAliquota.EOF = False Then
                                                        If TBAliquota!Simples = True Then
                                                            If TBFI!Estado = "EX" Then
                                                                'Quando é nota de importação os valores de PIS e Cofins já estão inclusos nas despesas acessorias
                                                                Valor_PIS_Prod = 0
                                                                Valor_Cofins_Prod = 0
                                                            Else
                                                                Valor_PIS_Prod = Format(qt * (TBFI!Total_PIS_prod / TBFI!int_Qtd), "0.0000000000")
                                                                Valor_Cofins_Prod = Format(qt * (TBFI!Total_Cofins_prod / TBFI!int_Qtd), "0.0000000000")
                                                            End If
                                                            Valor_CSLL_Prod = Format(qt * (TBFI!Total_CSLL_prod / TBFI!int_Qtd), "0.0000000000")
                                                            Valor_IRPJ_Prod = Format(qt * (TBFI!Total_IRPJ_prod / TBFI!int_Qtd), "0.0000000000")
                                                            'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário - Valor desc.) + (Valor ICMS + Valor do frete + Valor Seguro + Valor despesas + Valor PIS + Valor Cofins + Valor CSLL + Valor IRPJ)
                                                            Qtd = Format(qt * (IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario) - QuantsolicitadoN2) + (Valor1 + Valor2 + ValorPagar + ValorPago + Valor_PIS_Prod + Valor_Cofins_Prod + Valor_CSLL_Prod + Valor_IRPJ_Prod), "0.0000000000")
                                                        ElseIf TBAbrir!Real = True Then
                                                                Valor_PIS_Prod = Format(qt * (TBFI!Total_PIS_prod / TBFI!int_Qtd), "0.0000000000")
                                                                Valor_Cofins_Prod = Format(qt * (TBFI!Total_Cofins_prod / TBFI!int_Qtd), "0.0000000000")
                                                                'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário + Valor do frete + Valor Seguro + Valor despesas) - (Valor desc. + Valor ICMS + Valor PIS + Valor Cofins)
                                                                Qtd = (qt * IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario)) + Valor2 + ValorPagar + ValorPago
                                                                Qtd = Format(Qtd - (QuantsolicitadoN2 + Valor1 + Valor_PIS_Prod + Valor_Cofins_Prod), "0.0000000000")
                                                            Else
                                                                'VALOR UNITÁRIO DO ESTOQUE = (Valor unitário + Valor do frete + Valor Seguro + Valor despesas) - (Valor desc. + Valor ICMS)
                                                                Qtd = (qt * IIf(IsNull(TBFI!dbl_ValorUnitario), "0", TBFI!dbl_ValorUnitario)) + Valor2 + ValorPagar + ValorPago
                                                                Qtd = Format(Qtd - (QuantsolicitadoN2 + Valor1), "0.0000000000")
                                                            End If
                                                    End If
                                                Else
                                                    Set TBPedido = CreateObject("adodb.recordset")
                                                    TBPedido.Open "Select CPL.Quant_Comp, CPL.preco_unitario_desconto, CPL.vlrICMS from Compras_pedido_lista CPL INNER JOIN Compras_comercial CC ON CC.IDpedido = CPL.IDpedido where CPL.IdLista = " & TBFIltro!IDlista & " and CC.Moeda = 'REAL'", Conexao, adOpenKeyset, adLockOptimistic
                                                    If TBPedido.EOF = False Then
                                                        If TBPedido!Quant_Comp <> 0 Then valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto) - (IIf(IsNull(TBPedido!vlrICMS), "0", TBPedido!vlrICMS) / IIf(IsNull(TBPedido!Quant_Comp), "0", TBPedido!Quant_Comp)) Else valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
                                                    End If
                                                    TBPedido.Close
                                                End If
                                                TBFI.Close
                                            End If
                                            TBFIltro.Close
                            End If
                            
                            TBAbrir!Familia = TBEstoque!Classe
                            TBEstoque!valor_unitario = Format(valor, "###,##0.0000000000")
                            TBAbrir!VlrUnit = Format(valor, "###,##0.0000000000")
                            If IsNull(TBAbrir!Entrada) = False And TBAbrir!Entrada <> "0" Then TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Entrada), 0, TBAbrir!Entrada) Else TBAbrir!vlrTotal = valor * IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
                            TBAbrir.Update
                            TBAbrir.MoveNext
                        Loop
                    End If
                    
                    Total = Format(Entrada - Saida, "###,##0.0000000000")
                    TotalPC = Format(EntradaPC - SaidaPC, "###,##0.0000000000")
                    
                    If TBEstoque!local_armaz = "" Or IsNull(TBEstoque!local_armaz) = True Then TBEstoque!local_armaz = "N/A"
                    If ControlaEstoque = True Then
                        TBEstoque!estoque_venda = Total
                        TBEstoque!estoque_real = Total
                        TBEstoque!estoque_real_PC = TotalPC
                        TBEstoque!Valor_total = Format(valor * TBEstoque!estoque_real, "###,##0.00")
                    Else
                        TBEstoque!estoque_venda = 0
                        TBEstoque!estoque_real = 0
                        TBEstoque!estoque_real_PC = 0
                        TBEstoque!Valor_total = 0
                    End If
                    TBEstoque.Update
                    
                    Entrada = 0
                    EntradaPC = 0
                    Saida = 0
                    SaidaPC = 0
                    Total = 0
                    TotalPC = 0
                    valor = 0
                    TBEstoque.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBEstoque.Close
        End If
        
        If .Chk4.Value = 1 Then
            'Verifica pedidos de compra com centro de custo e produto que controla estoque
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select CP.ID_empresa, CPL.Desenho, CPL.Descricao, CPL.Familia, CPL.UN, CPL.Unidade_com, P.peso_metro, P.un_kg, EC.IDestoque, EC.Estoque_real from (((Compras_pedido_lista CPL INNER JOIN Compras_pedido_lista_custo CPLC ON CPLC.IDLista = CPL.IdLista) INNER JOIN projproduto P ON P.Desenho = CPL.Desenho) INNER JOIN Compras_pedido CP ON CP.IDpedido = CPL.IDpedido) INNER JOIN Estoque_Controle EC ON EC.Lote = CP.Pedido and EC.Desenho = CPL.Desenho where P.Estoque = 'True' and EC.Estoque_real > 0 and CPL.Tipo = 'P' and CPLC.ID_CC IS NOT NULL and CPLC.ID_CC <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBPedido.RecordCount
                PBLista.Value = 1
                Contador = 0
                
                'Cria requisição
                Set TBCompras = CreateObject("adodb.recordset")
                TBCompras.Open "Select * from Requisicao_materiais", Conexao, adOpenKeyset, adLockOptimistic
                TBCompras.AddNew
                TBCompras!ID_empresa = TBPedido!ID_empresa
                TBCompras!Responsavel = "PROCAM"
                TBCompras!Data = Date
                TBCompras!status = "RETIRADA"
                TBCompras!DtValidacao = Now
                TBCompras!RespValidacao = "PROCAM"
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from Requisicao_materiais where Year(data) = '" & Year(Date) & "' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Numero = Left(TBAbrir!requisicao, Len(TBAbrir!requisicao) - 3)
                    Numero = Right(Numero, 5) + 1
                Else
                    Numero = 1
                End If
                TBAbrir.Close
                a = "RM-" & FunTamanhoTextoZeroEsq(Numero, 5) & "/" & Right(Year(Date), 2)
                TBCompras!requisicao = a
                TBCompras!Obs = "ACERTO DO ESTOQUE *** PRODUTO QUE CONTROLA ESTOQUE E TEM CENTRO DE CUSTO NO PEDIDO DE COMPRA ***"
                TBCompras.Update
                ID_RM = TBCompras!ID
                
                Do While TBPedido.EOF = False
                    'Salva o produto na RM
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from Requisicao_materiais_lista", Conexao, adOpenKeyset, adLockOptimistic
                    TBCompras.AddNew
                    TBCompras!IDRequisicao = ID_RM
                    TBCompras!Data = Date
                    TBCompras!Responsavel = "PROCAM"
                    TBCompras!status = "RETIRADO"
                    TBCompras!Desenho = TBPedido!Desenho
                    TBCompras!Quant = TBPedido!estoque_real
                    TBCompras!quant_saida = TBPedido!estoque_real
                    TBCompras!Familia = TBPedido!Familia
                    TBCompras!Descricao = TBPedido!Descricao
                    TBCompras!Un = TBPedido!Un
                    TBCompras!Unidade_com = TBPedido!Unidade_com
                    TBCompras!ID_CC = Null
                    TBCompras!Data_autorizacao = Null
                    TBCompras!Autorizado = ""
                    TBCompras!Obs = Null
                    TBCompras.Update
                    TBCompras.Close
                
                    'Retira o produto do estoque
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select * from estoque_controle where IDestoque = " & TBPedido!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBEstoque.EOF = False Then
                        qtdeliberada = 0
                        qtdeliberadaPC = 0
                        qtdeliberar = 0
                        qtdeliberarPC = 0
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select Sum(Entrada) as qtdeliberada, Sum(ISNULL(Entrada_PC, 0)) as qtdeliberadaPC, Sum(Saida) as qtdeliberar, Sum(ISNULL(Saida_PC, 0)) as qtdeliberarPC from Estoque_movimentacao where IDestoque = " & TBPedido!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            qtdeliberada = IIf(IsNull(TBFI!qtdeliberada), 0, TBFI!qtdeliberada)
                            'qtdeliberadaPC = IIf(IsNull(TBFI!qtdeliberadaPC), 0, TBFI!qtdeliberadaPC)
                            qtdeliberar = IIf(IsNull(TBFI!qtdeliberar), 0, TBFI!qtdeliberar)
                            'qtdeliberarPC = IIf(IsNull(TBFI!qtdeliberarPC), 0, TBFI!qtdeliberarPC)
                            QtdeEstoque = Format(qtdeliberada - (qtdeliberar + TBPedido!estoque_real), "###,##0.0000")
                            'QtdeEstoquePC = Format(qtdeliberadaPC - (qtdeliberarPC + QtdeSaidaPC), "###,##0.0000")
                        End If
                        TBFI.Close
                        
                        TBEstoque!peso_unit = TBPedido!peso_metro
                        'TBEstoque!Pedido = IIf(txtPedidoCompra = "", Null, txtPedidoCompra)
                       
                        TBEstoque!estoque_real = QtdeEstoque
                        'TBEstoque!estoque_real_PC = QtdeEstoquePC
                        TBEstoque!estoque_venda = QtdeEstoque
                        TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * QtdeEstoque, "###,##0.00")
                               
                        Set TBProduto = CreateObject("adodb.recordset")
                        TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
                        TBProduto.AddNew
                        TBProduto!Operacao = "SAIDA_REQUISICAO"
                        TBProduto!Documento = a
                        TBProduto!LOTE = TBEstoque!LOTE
                        TBProduto!Desenho = TBEstoque!Desenho
                        TBProduto!Data = Date
                        TBProduto!Descricao = TBEstoque!Descricao
                        TBProduto!Familia = TBEstoque!Classe
                        TBProduto!Requisitante = "PROCAM"
                        TBProduto!Responsavel = "PROCAM"
                        TBProduto!IDEstoque = TBEstoque!IDEstoque
                        TBProduto!OE = a
                        TBProduto!Destino = "Interno"
                        TBProduto!Terceiros = False
                        
                        TBProduto!Saida = TBPedido!estoque_real
                        'TBProduto!Saida_PC = IIf(txtquantretirado_PC = "", 0, txtquantretirado_PC)
                        TBProduto!estoque_venda = QtdeEstoque
                    
                        'Atualiza valor do material no estoque
                        TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
                        TBProduto!vlrTotal = Format(TBPedido!estoque_real * TBProduto!VlrUnit, "###,##0.00")
                    
                        TBEstoque.Update
                        TBProduto.Update
                        TBProduto.Close
                    End If
                    TBEstoque.Close
                    TBPedido.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBPedido.Close
        End If
        
        If .Chk5.Value = 1 Then
            'Corrige empenho no RE
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select Ordem, Ordemempenho, Qtde_empenho from Producao_pedidos where Ordemempenho IS NOT NULL and Ordemempenho <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                PBLista.Min = 0
                PBLista.Max = TBPedido.RecordCount
                PBLista.Value = 1
                Contador = 0
                Do While TBPedido.EOF = False
                    valor = TBPedido!Qtde_empenho
                    
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select IDestoque, Desenho, Data, Responsavel from Estoque_Controle where Lote = '" & TBPedido!Ordem & "' and LEFT(status, 13) = 'ENTRADA_ORDEM' and Estoque_real > 0", Conexao, adOpenKeyset, adLockOptimistic
                    Do While TBEstoque.EOF = False
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ISNULL(Sum(Entrada), 0) as Valor1 from Estoque_movimentacao where IDestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Valor1 = TBAbrir!Valor1
                        End If
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ISNULL(Sum(Quantidade), 0) as Valor3 from Producao_NF_Consignada where Ordem = " & TBPedido!OrdemEmpenho & " and IDestoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Valor3 = TBAbrir!Valor3
                            valor = valor - Valor3
                        End If
                        TBAbrir.Close
                        If valor > 0 And Valor1 - Valor3 > 0 Then
                            Set TBGravar = CreateObject("adodb.recordset")
                            TBGravar.Open "Select * from Producao_NF_Consignada", Conexao, adOpenKeyset, adLockOptimistic
                            TBGravar.AddNew
                            TBGravar!Ordem = TBPedido!OrdemEmpenho
                            TBGravar!Codinterno = TBEstoque!Desenho
                            If valor <= Valor1 Then TBGravar!quantidade = valor Else TBGravar!quantidade = Valor1
                            TBGravar!IDEstoque = TBEstoque!IDEstoque
                            TBGravar!Data = TBEstoque!Data
                            TBGravar!Responsavel = TBEstoque!Responsavel
                            TBGravar!Qtde_saida = 0
                            TBGravar!Quantidade_PC = TBGravar!quantidade
                            TBGravar!Qtde_saida_PC = 0
                            TBGravar.Update
                            TBGravar.Close
                        End If
                        TBEstoque.MoveNext
                    Loop
                    TBEstoque.Close
                    TBPedido.MoveNext
                    Contador = Contador + 1
                    PBLista.Value = Contador
                Loop
            End If
            TBPedido.Close
        End If
        
        USMsgBox ("Atualização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Movimentação"
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
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Procentrada()
On Error GoTo tratar_erro
  
frmestoque_entrada.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcRetirada()
On Error GoTo tratar_erro
  
frmestoque_Retirar.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Label14_DblClick()
On Error GoTo tratar_erro

StrSql = "SELECT Data, Codigo, Descricao, Familia, Entrada, Saida, RE, Operacao from Estoque_Movimentacao_Diaria WHERE Operacao = 'SAIDA_NOTA' ORDER BY RE"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
Do While TBAbrir.EOF = False
 Conexao.Execute "update Estoque_movimentacao set data = '" & Format(TBAbrir!Data, "DD/MM/YYYY") & "' from Estoque_movimentacao where IDEstoque = " & TBAbrir!RE & " "
 TBAbrir.MoveNext
Loop
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub


ProcCarregaDadosLote

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaDadosFabricante()
On Error GoTo tratar_erro

ListaFabricante.ListItems.Clear
StrSql = "SELECT TOP (100) PERCENT PFAB.Part_number, FM.Fabricante FROM Fabricante_marca AS FM RIGHT OUTER JOIN Projproduto_fabricante AS PFAB ON FM.Id = PFAB.Idfabricante RIGHT OUTER JOIN Estoque_produtos AS EP ON PFAB.Codproduto = EP.codproduto WHERE EP.IdEstoque = '" & Lista.SelectedItem & "'"
Contador = 1
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
 If TBEstoque.EOF = False Then
     Do While TBEstoque.EOF = False
        With ListaFabricante.ListItems
            .Add , , Contador
            .Item(.Count).SubItems(1) = IIf(IsNull(TBEstoque!Part_number), "", TBEstoque!Part_number)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBEstoque!Fabricante), "", TBEstoque!Fabricante)
        End With
        Contador = Contador + 1
        TBEstoque.MoveNext
    Loop
 End If
 TBEstoque.Close
 


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaDadosLote()
On Error GoTo tratar_erro

TTE = 0
quantestoque = 0
quantestoquelote = 0
IDempresa = 0
RE = Lista.SelectedItem

Lista_Movimentacao.ListItems.Clear

'=================================================================================
'Corrige entradas por apontamento
'=================================================================================
'Set TBEstoque = CreateObject("adodb.recordset")
'TBEstoque.Open "Select * FROM estoque_movimentacao where idestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
'    If TBEstoque.EOF = False Then

'        Do While TBEstoque.EOF = False
'             If TBEstoque!IDApontamento <> Null Or TBEstoque!IDApontamento <> 0 Then
'                Set TBproducao = CreateObject("adodb.recordset")
'                TBproducao.Open "Select * FROM producaofases where idproducao = " & TBEstoque!IDApontamento, Conexao, adOpenKeyset, adLockOptimistic
'                If TBproducao.EOF = True Then
'                StrSql = "delete from estoque_movimentacao Where IDApontamento = '" & TBEstoque!IDApontamento & "'"
'                Conexao.Execute StrSql
'                End If
'                TBproducao.Close
'            End If
'        TBEstoque.MoveNext
'        Loop
'    End If
'TBEstoque.Close

'=================================================================================
'Dados de estoque do RE (Registro do estoque)
'=================================================================================
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM estoque_Controle_saldo_RE where idestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select sum(entrada) - Sum(saida) As EstoqueSaldo FROM estoque_movimentacao where idestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
       EstoqueSaldo = TBEstoque!EstoqueSaldo
    End If
    txtLote.Text = TBAbrir!LOTE
 
    NomeCampo = "o local de armazenamento"
    If IsNull(TBAbrir!local_armaz) = False And TBAbrir!local_armaz <> "" Then txtLocal_armaz = TBAbrir!local_armaz
    NomeCampo = "a família"
    If IsNull(TBAbrir!Classe) = False And TBAbrir!Classe <> "" Then cmbfamilia = TBAbrir!Classe
1:
    txtlocalização.Text = Lista.SelectedItem
    Txt_cod_ref = IIf(IsNull(TBAbrir!Ref), "", TBAbrir!Ref)
    Txt_n_serie = IIf(IsNull(TBAbrir!Numero_serie), "", TBAbrir!Numero_serie)
    txtvencimento.Value = Format(IIf(IsNull(TBAbrir!Vencimento), Date, TBAbrir!Vencimento), "dd/mm/yyyy")
    Txt_qtde_estoqueRE = Format(EstoqueSaldo, "###,##0.0000")
    
'====================================================
' Verifica empenho materia prima na ordem de produção
'====================================================
StrSql = "select idestoque,codinterno, SUM(Quantidade-qtde_saida) AS Empenhado from Producao_NF_Consignada Where IDestoque =" & Lista.SelectedItem & " GROUP BY IDestoque, codinterno Having (SUM(Quantidade-qtde_saida))> 0  order by IDestoque"
'Debug.print StrSql

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    EmpProd = IIf(IsNull(TBFI!Empenhado), 0, TBFI!Empenhado)
Else
    EmpProd = 0
End If
TBFI.Close
'====================================================
' Verifica empenho Produto na venda
'====================================================
StrSql = "Select ID_estoque, SUM(Qtde_empenhada-Qtde_saida) As empenhado from Estoque_Controle_Empenho_Vendas Where ID_estoque =" & Lista.SelectedItem & " group by ID_estoque Having SUM(Qtde_empenhada-Qtde_saida)>0 order by ID_estoque"
'Debug.print StrSql

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Empvend = IIf(IsNull(TBFI!Empenhado), 0, TBFI!Empenhado)
Else
    Empvend = 0
End If
TBFI.Close


Empenhado = Empvend + EmpProd

    Txt_qtde_empenhoRE = Format(Empenhado, "###,##0.0000")
    Txt_qtde_est_dispRE = Format(EstoqueSaldo - Empenhado, "###,##0.0000")  'Format(TBAbrir!Estoque_disponivel, "###,##0.0000")
    'Txt_qtde_est_disp_PCRE = Format(TBAbrir!Estoque_disponivel_PC, "###,##0.0000")
    
    Qtde = 0
    Qtd = 0
    
'=====================================================
' Busca quantidades em estoque de terceiros
'======================================================
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Saida) as qtde from estoque_movimentacao EM INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = EM.ID_prod_NF where EM.idestoque = " & Lista.SelectedItem & " and EM.destino = 'Terceiros' and NFP.Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtde = IIf(IsNull(TBFI!Qtde), 0, TBFI!Qtde)
    End If
    
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Entrada) as qtd from ((estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = EM.Documento and NF.txt_Razao_Nome = EC.Fornecedor) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID and NFP.int_Cod_Produto = EM.Desenho where EM.idestoque = " & Lista.SelectedItem & " and EM.pedidocompra IS NOT NULL and NFP.Retorno = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    
    If TBFI.EOF = False Then
        Qtd = IIf(IsNull(TBFI!Qtd), 0, TBFI!Qtd)
    End If
    
    Txt_qtde_est_terc = Format(Qtde - Qtd, "###,##0.0000")
    
    Txt_valor_unitRE = Format(TBAbrir!valor_unitario, "###,##0.000000")
    Txt_valor_total_estRE = Format((TBAbrir!valor_unitario * EstoqueSaldo), "###,##0.000000")

'=================================================================================
'Dados de estoque do item
'=================================================================================
    Set TBFI = CreateObject("adodb.recordset")
    StrSql = "Select Sum(Entrada)-Sum(Saida) as Saldo, avg(VlrUnit) as vlrUnitMedio from estoque_movimentacao Where Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and ID_Empresa = '" & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & "'  and Data <= '" & Date & "' group By Desenho"
    'Debug.print StrSql
    
    TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    
    If TBFI.EOF = False Then
        Txt_qtde_estoque = Format(TBFI!Saldo, "###,##0.0000")
        SaldoItem = Txt_qtde_estoque
    End If
    
      '  Txt_custo_medio_unit = Format(TBFI!vlrUnitMedio, "###,##0.000000")

' Verifica empenho item na ordem de produção
    StrSql = "select codinterno, SUM(Quantidade-qtde_saida) AS Empenhado from Producao_NF_Consignada where codinterno = '" & Lista.SelectedItem.ListSubItems(3) & "' GROUP BY codinterno Having (SUM(Quantidade-qtde_saida))> 0  order by Codinterno"
'Debug.print StrSql

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    EmpItemProd = IIf(IsNull(TBFI!Empenhado), 0, TBFI!Empenhado)
Else
    EmpItemProd = 0
End If
TBFI.Close

'====================================================
' Verifica empenho Produto na venda
'====================================================
StrSql = "select EC.Desenho ,SUM(Qtde_empenhada-Qtde_saida) AS EMPENHADO from Estoque_Controle_Empenho_Vendas ECEV inner join Estoque_Controle EC on ECEV.ID_estoque = EC.IdEstoque Where EC.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' Group by ec.Desenho  Having SUM(Qtde_empenhada-Qtde_saida)>0"
'Debug.print StrSql

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    EmpItemvend = IIf(IsNull(TBFI!Empenhado), 0, TBFI!Empenhado)
Else
    EmpItemvend = 0
End If
TBFI.Close


    Empenhado = EmpItemvend + EmpItemProd


    Txt_qtde_empenho = Format(Empenhado, "###,##0.0000")
    '====================================================================
    ' Carrega dados do item disponivel - empenho
    '====================================================================
    
    If SaldoItem > 0 Or IsNull(SaldoItem) = False Then
    Txt_qtde_est_disp = Format(SaldoItem - Empenhado, "###,##0.0000")
    Else
    
    Txt_qtde_est_disp = Format("0", "###,##0.0000")
    
    End If
    'Txt_qtde_est_disp_PC = Format(quantnovo, "###,##0.0000")
        
    Qtde = 0
    Qtd = 0
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Saida) as qtde from estoque_movimentacao EM INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = EM.ID_prod_NF where EM.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and EM.destino = 'Terceiros' and NFP.Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtde = IIf(IsNull(TBFI!Qtde), 0, TBFI!Qtde)
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Entrada) as qtd from ((estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque) INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.int_NotaFiscal = EM.Documento and NF.txt_Razao_Nome = EC.Fornecedor) INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID and NFP.int_Cod_Produto = EM.Desenho where EM.Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and EM.pedidocompra IS NOT NULL and NFP.Retorno = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtd = IIf(IsNull(TBFI!Qtd), 0, TBFI!Qtd)
    End If
    Txt_qtde_est_terc = Format(Qtde - Qtd, "###,##0.0000")
    
    If Txt_qtde_est_disp <> 0 Then
    CTMedioEst = Txt_valor_total_est / Txt_qtde_est_disp
    Txt_custo_medio_unit = Format(CTMedioEst, "###,##0.000000")
    End If
    
    ProcCarregaListaMovimentacao
    CodigoLista = Lista.SelectedItem.index
    
    txtFornecedor.Text = Lista.SelectedItem.ListSubItems.Item(12).Text
    txtLocalArmazenamento = Lista.SelectedItem.ListSubItems.Item(8).Text
    txtfamilia.Text = Lista.SelectedItem.ListSubItems.Item(7).Text
    
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(EM.Saida) as TTsaida, Sum(EM.entrada) as TTentrada  from estoque_movimentacao EM where IDEstoque = '" & txtlocalização.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
       txtTTEntrada.Text = Format(TBFI!ttEntrada, "###,##0.0000")
       txtTTSaida.Text = Format(TBFI!ttsaida, "###,##0.0000")
    End If

End If

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        USMsgBox ("Não foi encontrado " & NomeCampo & " deste registro."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaListaMovimentacao()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub

Lista_Movimentacao.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from estoque_movimentacao where idestoque = " & Lista.SelectedItem & " order by Data desc, Idoperacao desc", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBAbrir.EOF = False
        With Lista_Movimentacao.ListItems
            .Add , , TBAbrir!IDoperacao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Operacao), "", TBAbrir!Operacao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Entrada), "0,0000", Format(TBAbrir!Entrada, "###,##0.0000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Entrada_PC), "0,0000", Format(TBAbrir!Entrada_PC, "###,##0.0000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Saida), "0,0000", Format(TBAbrir!Saida, "###,##0.0000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!Saida_PC), "", Format(TBAbrir!Saida_PC, "###,##0.0000"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!Documento), "", TBAbrir!Documento)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBAbrir!Requisitante), "", TBAbrir!Requisitante)
            If TBAbrir!Destino = "Terceiros" Then .Item(.Count).SubItems(11) = "Terceiros (Remessa forn.)" Else .Item(.Count).SubItems(11) = "Interno/Cliente"
            
            If TBAbrir!Entrada > 0 Or TBAbrir!Destino = "Terceiros" Then
                .Item(.Count).SubItems(12) = IIf(IsNull(TBAbrir!Pedidocompra), "", TBAbrir!Pedidocompra)
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "Select Fornecedor from Compras_pedido where idpedido = " & IIf(IsNull(TBAbrir!IDpedido), 0, TBAbrir!IDpedido), Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then .Item(.Count).SubItems(13) = IIf(IsNull(TBPedido!Fornecedor), "", TBPedido!Fornecedor)
                TBPedido.Close
            Else
                If TBAbrir!Operacao = "SAIDA_NOTA" Or TBAbrir!Operacao = "SAIDA_NOTA_PARCIAL" Then
                    Set TBControleNF = CreateObject("adodb.recordset")
                    TBControleNF.Open "Select TDNF.txt_Razao_Nome from tbl_Detalhes_Nota TDN INNER JOIN tbl_Dados_Nota_Fiscal TDNF ON TDN.ID_Nota = TDNF.ID where TDN.Int_codigo = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF), Conexao, adOpenKeyset, adLockOptimistic
                    If TBControleNF.EOF = False Then
                        Set TBPedido = CreateObject("adodb.recordset")
                        TBPedido.Open "Select VP.Ncotacao from (tbl_Detalhes_Nota_pedidos TDNP INNER JOIN vendas_carteira VC ON TDNP.ID_carteira = VC.Codigo) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.cotacao where TDNP.ID_prod_NF = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF), Conexao, adOpenKeyset, adLockOptimistic
                        If TBPedido.EOF = False Then .Item(.Count).SubItems(12) = IIf(IsNull(TBPedido!Ncotacao), "", TBPedido!Ncotacao)
                        TBPedido.Close
                        
                        .Item(.Count).SubItems(13) = IIf(IsNull(TBControleNF!txt_Razao_Nome), "", TBControleNF!txt_Razao_Nome)
                    End If
                    TBControleNF.Close
                End If
            End If
            .Item(.Count).SubItems(14) = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
        End With
        TBAbrir.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_movimentacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_Movimentacao
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Status_movimentacao = .ListItems(InitFor).SubItems(2)
                If Status_movimentacao = "SAIDA_ALMOXARIFADO" Or Status_movimentacao = "ENTRADA_ALMOXARIFADO" Or Status_movimentacao = "DEVOLUCAO_ALMOXARIFADO C/ PROB." Or Status_movimentacao = "SAIDA_SUCATA" Or Status_movimentacao = "ENTRADA_SUCATA" Or Status_movimentacao = "SAIDA_RETALHO" Or Status_movimentacao = "ENTRADA_RETALHO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL" Then
                    GoTo Proximo
                ElseIf Status_movimentacao = "ENTRADA_INVENTÁRIO" Or Status_movimentacao = "SAIDA_INVENTÁRIO" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from estoque_movimentacao where idoperacao = " & .ListItems(InitFor) & " and id_inventario <> 0 and ID_inventario IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        TBAbrir.Close
                        GoTo Proximo
                    End If
                    TBAbrir.Close
                Else
                    'Verifica se a entrada esta vinculada a ordem
                    If Left(Status_movimentacao, 7) = "ENTRADA" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select PNC.Ordem from estoque_movimentacao EM INNER JOIN Producao_NF_Consignada PNC ON EM.IDestoque = PNC.Idestoque where EM.idoperacao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            'Verifica qtde. de entrada do RE
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select ROUND(SUM(ISNULL(Entrada, 0)), 3) as Valor from estoque_movimentacao where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                            End If
                            'Verifica qtde. empenhada do RE
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select ROUND(SUM(ISNULL(Quantidade, 0)), 3) as Valor1 from Producao_NF_Consignada where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                            End If
                            Permitido = True
                            If valor - Valor1 <= 0 Then
                                Permitido = False
                            ElseIf (valor - .ListItems(InitFor).SubItems(4)) - Valor1 < 0 Then
                                Permitido = False
                            End If
                            If Permitido = False Then
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                        End If
                        TBAbrir.Close
                    End If
                    
                    'Verifica se o resultado da ordem esta validado
                    If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Or Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select P.Ordem from estoque_movimentacao EM INNER JOIN Producao P ON EM.Documento = P.Ordem where EM.idoperacao = " & .ListItems(InitFor) & " and P.RespValidacao_Custo IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            TBAbrir.Close
                            GoTo Proximo
                        End If
                        TBAbrir.Close
                    End If
                    
                    'Verifica se a RE tem movimentações mais recentes
                    If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Or Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                        If Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                            Set TBAfericao = CreateObject("adodb.recordset")
                            TBAfericao.Open "SELECT IdEstoque, idoperacao FROM estoque_movimentacao WHERE IdTrocaLocal = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                            If TBAfericao.EOF = False Then
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & TBAfericao!IDEstoque & " AND Idoperacao <> " & TBAfericao!IDoperacao & " AND idoperacao > " & TBAfericao!IDoperacao, Conexao, adOpenKeyset, adLockReadOnly
                                If TBAbrir.EOF = False Then
                                    TBAbrir.Close
                                    TBAfericao.Close
                                    GoTo Proximo
                                End If
                                TBAbrir.Close
                            End If
                            TBAfericao.Close
                        Else
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & Lista.SelectedItem & " AND Idoperacao <> " & .ListItems(InitFor) & " AND  idoperacao > " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                            If TBAbrir.EOF = False Then
                                TBAbrir.Close
                                GoTo Proximo
                            End If
                            TBAbrir.Close
                        End If
                    End If
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_Movimentacao, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcAtualizaStatus_RM()
On Error GoTo tratar_erro

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Requisicao_materiais where requisicao = '" & TBAbrir!Documento & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'REQUISIT.'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = True Then
            TBproducao!status = "ABERTA"
        Else
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'RETIRADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = True Then
                TBproducao!status = "RETIRADA"
            Else
                TBproducao!status = "PARCIAL"
            End If
        End If
        TBproducao.Update
    End If
    TBMateriaprima.Close
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcOrdem()
On Error GoTo tratar_erro

If Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from projproduto where Desenho = '" & Lista.SelectedItem.ListSubItems(3) & "' and SubTipoItem <> 0 and SubTipoItem <> 4", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        If TBCiclo!QuantProd - Qtd <= 0 Then TBCiclo!QuantProd = 0 Else TBCiclo!QuantProd = TBCiclo!QuantProd - Qtd
    End If
    TBProduto.Close
    If TBCiclo!QuantProd <> 0 Then
        TBCiclo!CPR = IIf(IsNull(TBCiclo!CTTReal), 0, TBCiclo!CTTReal) / TBCiclo!QuantProd
    Else
        TBCiclo!CPR = 0
        TBCiclo!Controlado_estoque = False
    End If
    If TBCiclo!QuantProd < TBCiclo!Quant Then
        TBCiclo!DataEntrega = Null
        TBCiclo!Concluida = False
        TBCiclo!pronta = "NÃO"
        If TBCiclo!status <> "Entregue" Then TBCiclo!status = "Aberta"
        
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from producao where Ordem = " & TBCiclo!Ordem & " and Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            NomeTabelaAp = "ProducaoFases_Backup"
        Else
            NomeTabelaAp = "ProducaoFases"
        End If
        
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from ordemservico where Ordem = " & TBCiclo!Ordem & " and pronto = 'SIM'", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            Do While TBproducao.EOF = False
                TBproducao!Pronto = "NÃO"
                TBproducao!DataConclusao = Null
                TBproducao!status = Null
                TBproducao.Update
                'Filtra todos os eventos desta OS na tabela producaofases para marcar como fase pronta
                Conexao.Execute "Update " & NomeTabelaAp & " Set pronto = 'NÃO' where idfase = " & TBproducao!IDProducao
                TBproducao.MoveNext
            Loop
        End If
        TBproducao.Close
    End If
    TBCiclo.Update
End If
'==================================
Modulo = "Estoque/Movimentação/Entrada"
Evento = "Alterar OF p/ não concluída"
ID_documento = TBCiclo!NOF
Documento = "Ordem: " & TBCiclo!Ordem
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_movimentacao_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_Movimentacao
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            Status_movimentacao = .ListItems(InitFor).SubItems(2)
            If Status_movimentacao = "SAIDA_ALMOXARIFADO" Or Status_movimentacao = "ENTRADA_ALMOXARIFADO" Or Status_movimentacao = "DEVOLUCAO_ALMOXARIFADO C/ PROB." Or Status_movimentacao = "SAIDA_SUCATA" Or Status_movimentacao = "ENTRADA_SUCATA" Or Status_movimentacao = "SAIDA_RETALHO" Or Status_movimentacao = "ENTRADA_RETALHO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL_CONSIGNAÇÃO" Or Status_movimentacao = "ENTRADA_NOTA_FISCAL" Then
                If Status_movimentacao = "SAIDA_SUCATA" Or Status_movimentacao = "SAIDA_RETALHO" Then
                    USMsgBox ("Não é permitido excluir esta movimentação."), vbExclamation, "CAPRIND v5.0"
                ElseIf Status_movimentacao = "ENTRADA_SUCATA" Or Status_movimentacao = "ENTRADA_RETALHO" Then
                        USMsgBox ("Só é permitido excluir o lote desta movimentação, utilizando o botão (Excluir sucata/ret.)."), vbExclamation, "CAPRIND v5.0"
                    Else
                        USMsgBox ("Não é permitido excluir este tipo de movimentação neste módulo."), vbExclamation, "CAPRIND v5.0"
                End If
                .ListItems.Item(InitFor).Checked = False
            ElseIf Status_movimentacao = "ENTRADA_INVENTÁRIO" Or Status_movimentacao = "SAIDA_INVENTÁRIO" Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select idoperacao from estoque_movimentacao where idoperacao = " & .ListItems(InitFor) & " and id_inventario <> 0 and ID_inventario IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    USMsgBox ("Não é permitido excluir este tipo de movimentação neste módulo."), vbExclamation, "CAPRIND v5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                TBAbrir.Close
            Else
                'Verifica se a entrada esta vinculada a ordem
                If Left(Status_movimentacao, 7) = "ENTRADA" And Status_movimentacao <> "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select PNC.Ordem from estoque_movimentacao EM INNER JOIN Producao_NF_Consignada PNC ON EM.IDestoque = PNC.Idestoque where EM.idoperacao = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        'Verifica qtde. de entrada do RE
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ROUND(SUM(ISNULL(Entrada, 0)), 3) as Valor from estoque_movimentacao where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
                        End If
                        'Verifica qtde. empenhada do RE
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select ROUND(SUM(ISNULL(Quantidade, 0)), 3) as Valor1 from Producao_NF_Consignada where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
                        End If
                        Permitido = True
                        If valor - Valor1 <= 0 Then
                            Permitido = False
                        ElseIf (valor - .ListItems(InitFor).SubItems(4)) - Valor1 < 0 Then
                            Permitido = False
                        End If
                        If Permitido = False Then
                            OPTexto = ""
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select Ordem from Producao_NF_Consignada where IDestoque = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                Do While TBAbrir.EOF = False
                                    If OPTexto = "" Then OPTexto = TBAbrir!Ordem Else OPTexto = OPTexto & " | " & TBAbrir!Ordem
                                    TBAbrir.MoveNext
                                Loop
                            End If
                            USMsgBox ("Não é permitido excluir esta movimentação, pois a mesma está sendo vinculada a(s) ordem(ns): " & vbCrLf & OPTexto), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            Exit Sub
                        End If
                    End If
                    TBAbrir.Close
                End If
                
                'Verifica se o resultado da ordem esta validado
                If Status_movimentacao = "ENTRADA_ORDEM" Or Status_movimentacao = "ENTRADA_ORDEM_PARCIAL" Or Status_movimentacao = "SAIDA_ORDEM" Or Status_movimentacao = "SAIDA_ORDEM_PARCIAL" Then
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select P.Ordem from estoque_movimentacao EM INNER JOIN Producao P ON EM.Documento = P.Ordem where EM.idoperacao = " & .ListItems(InitFor) & " and P.RespValidacao_Custo IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        USMsgBox ("Não é permitido excluir esta movimentação, pois o resultado da ordem " & TBAbrir!Ordem & " já foi validado."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                        TBAbrir.Close
                        Exit Sub
                    End If
                    TBAbrir.Close
                End If
                
                'Verifica se a RE tem movimentações mais recentes
                If Status_movimentacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO" Or Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                    If Status_movimentacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO" Then
                        Set TBAfericao = CreateObject("adodb.recordset")
                        TBAfericao.Open "SELECT IdEstoque, idoperacao FROM estoque_movimentacao WHERE IdTrocaLocal = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                        If TBAfericao.EOF = False Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & TBAfericao!IDEstoque & " AND Idoperacao <> " & TBAfericao!IDoperacao & " AND idoperacao > " & TBAfericao!IDoperacao, Conexao, adOpenKeyset, adLockReadOnly
                            If TBAbrir.EOF = False Then
                                USMsgBox ("Não é permitido excluir esta movimentação, pois exitem movimentações mais recentes no RE " & TBAbrir!IDEstoque & "."), vbExclamation, "CAPRIND v5.0"
                                .ListItems.Item(InitFor).Checked = False
                                TBAbrir.Close
                                TBAfericao.Close
                                Exit Sub
                            End If
                            TBAbrir.Close
                        End If
                        TBAfericao.Close
                    Else
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "SELECT idestoque FROM estoque_movimentacao WHERE idestoque = " & Lista.SelectedItem & " AND Idoperacao <> " & .ListItems(InitFor) & " AND  idoperacao > " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockReadOnly
                        If TBAbrir.EOF = False Then
                            USMsgBox ("Não é permitido excluir esta movimentação, pois exitem movimentações mais recentes no RE."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            TBAbrir.Close
                            Exit Sub
                        End If
                        TBAbrir.Close
                    End If
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista_Movimentacao.ListItems.Clear
ProcLimpaCamposTotais

If cmbfiltrarpor = "RE" And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: Procentrada
    Case 3: ProcRetirada
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    Case 6: ProcStatus
    Case 7: ProcEstruturadoitem
    Case 8: ProcSucata
    Case 9: ProcExcluir_sucata
    Case 10: ProcLocalArmazenamento
    Case 11: ProcAlterar_valor
    Case 12: ProcCC
    Case 13: ProcAtualiza3
    'Case 15: ProcAjuda
    Case 16: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAtualiza3()
On Error GoTo tratar_erro

Set TBEstoque = CreateObject("adodb.recordset")
StrSql = "Select IDApontamento,entrada from Estoque_movimentacao where idapontamento is not null order by IDApontamento"
TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBEstoque.EOF = False Then
    Do While TBEstoque.EOF = False
    
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "Select quantidade from producaofases where idproducao = '" & TBEstoque!IDApontamento & "'"
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TBEstoque!Entrada = TBAbrir!quantidade
            TBEstoque.Update
        Else
        If TBEstoque!IDApontamento <> 0 Then
            StrSql = "delete from estoque_movimentacao Where IDApontamento = '" & TBEstoque!IDApontamento & "'"
            Conexao.Execute StrSql
            End If
        End If
    TBEstoque.MoveNext
Loop
End If
TBEstoque.Close

USMsgBox "Correção no estoque executada com sucesso!", vbInformation, "CAPRIND v5.0"


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAtualiza2()
On Error GoTo tratar_erro

'Set TBEstoque = CreateObject("adodb.recordset")
'StrSQL_OF = "SELECT RE, ValorUnitario from Estoque_Movimentacao_Diaria_Fiscal group by RE, valorUnitario"
'TBEstoque.Open StrSQL_OF, Conexao, adOpenKeyset, adLockOptimistic
'
'If TBEstoque.EOF = False Then
'Do While TBEstoque.EOF = False
'    StrSql = "UPDATE estoque_movimentacao SET vlrUnit = " & Replace(TBEstoque!Valorunitario, ",", ".") & " Where IDEstoque = " & TBEstoque!RE
'Debug.print StrSql
'
'Conexao.Execute StrSql
'TBEstoque.MoveNext
'
'Loop
'End If
'TBEstoque.Close

Set TBEstoque = CreateObject("adodb.recordset")
StrSQL_OF = "Select idoperacao,EM.IdEstoque, Data, Documento, vlrUnit, TDN.dbl_ValorUnitario from Estoque_movimentacao EM inner join tbl_Detalhes_Nota TDN on EM.Documento = TDN.Int_NotaFiscal and Em.Desenho = TDN.int_Cod_Produto Where EM.VlrUnit <> TDN.dbl_ValorUnitario"
TBEstoque.Open StrSQL_OF, Conexao, adOpenKeyset, adLockOptimistic

If TBEstoque.EOF = False Then
Do While TBEstoque.EOF = False
    StrSql = "UPDATE estoque_movimentacao SET vlrUnit = '" & Replace(TBEstoque!dbl_ValorUnitario, ",", ".") & "' Where IDestoque = '" & TBEstoque!IDEstoque & "'"
'Debug.print StrSql

Conexao.Execute StrSql
TBEstoque.MoveNext

Loop
End If
TBEstoque.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procVoltarEmpenhoLocal(idEstoqueEntrada As Long, idEstoqueSaida As Long)
On Error GoTo tratar_erro

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT ID_estoque, IdAntigoLocal, Qtde_empenhada FROM Estoque_Controle_Empenho_Vendas where id_estoque = " & idEstoqueEntrada, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False
        If IsNull(TBCFOP!IdAntigoLocal) = True Then
            TBCFOP!ID_estoque = idEstoqueSaida
            TBCFOP.Update
        Else
            Set TBCorretiva = CreateObject("adodb.recordset")
            TBCorretiva.Open "SELECT Qtde_empenhada FROM Estoque_Controle_Empenho_Vendas where id = " & TBCFOP!IdAntigoLocal, Conexao, adOpenKeyset, adLockOptimistic
            If TBCorretiva.EOF = False Then
                TBCorretiva!Qtde_empenhada = TBCFOP!Qtde_empenhada + TBCorretiva!Qtde_empenhada
                TBCorretiva.Update
            End If
            TBCorretiva.Close
            TBCFOP.Delete
        End If
        TBCFOP.MoveNext
    Loop
End If
TBCFOP.Close

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT IDestoque, IdAntigoLocal, Quantidade, Quantidade_PC FROM Producao_NF_Consignada where idestoque = " & idEstoqueEntrada, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Do While TBCFOP.EOF = False
        If IsNull(TBCFOP!IdAntigoLocal) = True Then
            TBCFOP!IDEstoque = idEstoqueSaida
            TBCFOP.Update
        Else
            Set TBCorretiva = CreateObject("adodb.recordset")
            TBCorretiva.Open "SELECT * FROM Producao_NF_Consignada where id = " & TBCFOP!IdAntigoLocal, Conexao, adOpenKeyset, adLockOptimistic
            If TBCorretiva.EOF = False Then
                TBCorretiva!quantidade = TBCFOP!quantidade + TBCorretiva!quantidade
                TBCorretiva!Quantidade_PC = TBCFOP!Quantidade_PC + TBCorretiva!Quantidade_PC
                TBCorretiva.Update
            End If
            TBCorretiva.Close
            TBCFOP.Delete
        End If
        TBCFOP.MoveNext
    Loop
End If
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
