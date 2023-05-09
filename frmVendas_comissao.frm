VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_comissao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Vendas - Relatórios - Comissão"
   ClientHeight    =   10035
   ClientLeft      =   300
   ClientTop       =   1590
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendas_comissao.frx":0000
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
   Begin VB.Frame Frame1 
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
      Height          =   825
      Left            =   55
      TabIndex        =   15
      Top             =   9180
      Width           =   15195
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13170
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Valor total da comissão."
         Top             =   375
         Width           =   1875
      End
      Begin VB.TextBox txtValorTotal_vendido 
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
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Valor total da comissão."
         Top             =   375
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total comissão :"
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
         Left            =   11550
         TabIndex        =   17
         Top             =   420
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total vendido :"
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
         Index           =   1
         Left            =   7470
         TabIndex        =   16
         Top             =   420
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6840
      Left            =   60
      TabIndex        =   5
      Top             =   2055
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12065
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Data venda"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "PI"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Nota fiscal"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cliente"
         Object.Width           =   6950
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Vendedor externo"
         Object.Width           =   5010
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Vendedor Interno"
         Object.Width           =   5010
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Comissão"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Valor comissão"
         Object.Width           =   1764
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   8910
      Width           =   11775
      _ExtentX        =   20770
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13980
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_comissao.frx":030A
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   19
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
         Name            =   "MS Sans Serif"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   1455
      Begin DrawSuite2022.USOptionButton optResumido 
         Height          =   315
         Left            =   180
         TabIndex        =   32
         ToolTipText     =   "Filtrar por período"
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         Caption         =   "Resumido"
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
         Value           =   -1  'True
      End
      Begin DrawSuite2022.USOptionButton optDetalhado 
         Height          =   315
         Left            =   180
         TabIndex        =   33
         ToolTipText     =   "Filtrar por período"
         Top             =   630
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         Caption         =   "Detalhado"
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
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções de filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   1530
      TabIndex        =   12
      Top             =   960
      Width           =   9855
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "frmVendas_comissao.frx":3103
         Left            =   7800
         List            =   "frmVendas_comissao.frx":3110
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Opções para filtro."
         Top             =   510
         Width           =   1905
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
         ItemData        =   "frmVendas_comissao.frx":312E
         Left            =   120
         List            =   "frmVendas_comissao.frx":3130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   510
         Width           =   1905
      End
      Begin VB.ComboBox cmbTexto 
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
         Height          =   315
         ItemData        =   "frmVendas_comissao.frx":3132
         Left            =   2040
         List            =   "frmVendas_comissao.frx":3134
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   510
         Width           =   5745
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   8520
         TabIndex        =   22
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label8 
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
         Left            =   645
         TabIndex        =   14
         Top             =   300
         Width           =   705
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
         Left            =   4170
         TabIndex        =   13
         Top             =   300
         Width           =   1470
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6840
      Left            =   60
      TabIndex        =   4
      Top             =   2055
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12065
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt. venda"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "PI"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Status"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Vendedor"
         Object.Width           =   5009
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Comissão"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valor comissão"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar por"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   11400
      TabIndex        =   28
      Top             =   960
      Width           =   1725
      Begin DrawSuite2022.USOptionButton optPeriodo 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         ToolTipText     =   "Filtrar por período"
         Top             =   630
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "Período"
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
      End
      Begin DrawSuite2022.USOptionButton optMesAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   30
         ToolTipText     =   "Filtrar por ano e mês"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Ano x Mês"
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
      End
      Begin DrawSuite2022.USOptionButton optAno 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   420
         TabIndex        =   31
         ToolTipText     =   "Filtrar por ano"
         Top             =   1110
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Ano"
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
      End
   End
   Begin VB.Frame FPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   13135
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   2115
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   690
         TabIndex        =   3
         ToolTipText     =   "Data final."
         Top             =   645
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
         Format          =   196149249
         CurrentDate     =   39799
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   690
         TabIndex        =   2
         ToolTipText     =   "Data inicio."
         Top             =   285
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
         Format          =   196149249
         CurrentDate     =   39799
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Left            =   240
         TabIndex        =   11
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         Left            =   300
         TabIndex        =   10
         Top             =   345
         Width           =   300
      End
   End
   Begin VB.Frame FAnoMes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ano e mês"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   13140
      TabIndex        =   23
      Top             =   960
      Width           =   2115
      Begin VB.ComboBox cmbdoMes 
         Height          =   330
         ItemData        =   "frmVendas_comissao.frx":3136
         Left            =   720
         List            =   "frmVendas_comissao.frx":3138
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Escolha o mês para filtrar"
         Top             =   660
         Width           =   945
      End
      Begin VB.ComboBox cmbdoAno 
         Height          =   330
         ItemData        =   "frmVendas_comissao.frx":313A
         Left            =   720
         List            =   "frmVendas_comissao.frx":313C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Escolha o ano para filtrar"
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   27
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do ano"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Label Lbl_relatorio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11910
      TabIndex        =   18
      Top             =   8940
      Width           =   2895
   End
End
Attribute VB_Name = "frmVendas_comissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaComboAno()
On Error GoTo tratar_erro
cmbdoAno.Clear
cmbdoMes.Clear

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
      If IsNull(TBAbrir!Ano) = False Then
        With cmbdoAno
            .AddItem IIf(IsNull(TBAbrir!Ano), "", TBAbrir!Ano)
        End With
       End If
        TBAbrir.MoveNext
    Loop
  End If
  TBAbrir.Close
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbdoAno_Change()
On Error GoTo tratar_erro

StrSql = "Select DISTINCT Mes from Vendas_relatorios_comissao_detalhado Where Ano = '" & cmbdoAno.Text & "'order by Mes"

procCarregaComboMes


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregaComboMes()
On Error GoTo tratar_erro
cmbdoMes.Clear
Dim Mes As String

If cmbdoAno.Text <> "" Then

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        Select Case TBAbrir!Mes
            Case 1: Mes = "Janeiro"
            Case 2: Mes = "Fevereiro"
            Case 3: Mes = "Março"
            Case 4: Mes = "Abril"
            Case 5: Mes = "Maio"
            Case 6: Mes = "Junho"
            Case 7: Mes = "Julho"
            Case 8: Mes = "Agosto"
            Case 9: Mes = "Setembro"
            Case 10: Mes = "Outubro"
            Case 11: Mes = "Novembro"
            Case 12: Mes = "Dezembro"
        End Select
        
        With cmbdoMes
            .AddItem Mes
        End With
        TBAbrir.MoveNext
    Loop
  End If
  TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbdoAno_Click()
On Error GoTo tratar_erro

StrSql = "Select DISTINCT Mes from Vendas_relatorios_comissao_detalhado Where Ano = '" & cmbdoAno.Text & "'order by Mes"

procCarregaComboMes


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    If Lista.ListItems.Count = 0 Then Exit Sub
    NomeRel = "Vendas_comissao_individual_detalhado.rpt"
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
    NomeRel = "Vendas_comissao_individual_resumido.rpt"
End If
                        
ProcImprimirRel FormulaRelatorio, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: Cmd_ajuda_Click
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaDetalhada()
On Error GoTo tratar_erro
Dim TotalComissao As Double
Dim TotalVendido As Double

Posicao = 0
Familiatext = ""
Contador1 = 1
Lista.ListItems.Clear
Lista1.ListItems.Clear

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBCarteira.EOF = False Then

    Posicao = TBCarteira.RecordCount
    
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
                With Lista.ListItems
                    .Add , , TBCarteira!CODIGO
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBCarteira!Datavendas), "", Format(TBCarteira!Datavendas, "dd/mm/yy"))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBCarteira!Ncotacao), "", TBCarteira!Ncotacao)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBCarteira!int_NotaFiscal), "", TBCarteira!int_NotaFiscal)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBCarteira!Liberacao), "", TBCarteira!Liberacao)
                    
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBCarteira!Vend_ext), "", TBCarteira!Vend_ext)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBCarteira!Comissao), 0, TBCarteira!Comissao & "%")
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBCarteira!ValorComissao), 0, Format(TBCarteira!ValorComissao, "###,##0.0000"))
         

                    .Item(.Count).SubItems(8) = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBCarteira!descricao_tecnica), "", TBCarteira!descricao_tecnica)
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBCarteira!quantidade), "", Format(TBCarteira!quantidade, "###,##0.0000"))
                    .Item(.Count).SubItems(11) = IIf(IsNull(TBCarteira!preco_unitario_desconto), "", (Format(TBCarteira!preco_unitario_desconto, "###,##0.0000")))
                    .Item(.Count).SubItems(12) = Format(IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote), "###,##0.00")
                    TotalComissao = TotalComissao + IIf(IsNull(TBCarteira!ValorComissao), 0, TBCarteira!ValorComissao)
                    TotalVendido = TotalVendido + Format(IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote), "###,##0.00")
                End With
        TBCarteira.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBCarteira.Close

txtValorTotal_vendido = Format(TotalVendido, "###,##0.00")
txtValorTotal = Format(TotalComissao, "###,##0.00")


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaResumida()
On Error GoTo tratar_erro
Dim TotalComissao As Double
Dim TotalVendido As Double

Posicao = 0
Familiatext = ""
Contador1 = 1
Lista.ListItems.Clear
Lista1.ListItems.Clear

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBCarteira.EOF = False Then
    Posicao = TBCarteira.RecordCount
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
                With Lista1.ListItems
                    .Add , , Contador 'TBCarteira!CODIGO
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBCarteira!Datavendas), "", Format(TBCarteira!Datavendas, "dd/mm/yy"))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBCarteira!Pedido), "", TBCarteira!Pedido)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBCarteira!Nota), "", TBCarteira!Nota)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBCarteira!status), "", TBCarteira!status)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBCarteira!Cliente), "", TBCarteira!Cliente)
                    
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBCarteira!Vend_ext), "", TBCarteira!Vend_ext)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBCarteira!vend_int), "", TBCarteira!vend_int)
                    .Item(.Count).SubItems(8) = Format(IIf(IsNull(TBCarteira!Totalpedido), 0, TBCarteira!Totalpedido), "###,##0.00")
                    
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBCarteira!Comissao), 0, TBCarteira!Comissao & "%")
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBCarteira!ValorComissao), 0, Format(TBCarteira!ValorComissao, "###,##0.0000"))
         
                    TotalComissao = TotalComissao + IIf(IsNull(TBCarteira!ValorComissao), 0, TBCarteira!ValorComissao)
                    TotalVendido = TotalVendido + Format(IIf(IsNull(TBCarteira!Totalpedido), 0, TBCarteira!Totalpedido), "###,##0.00")
                End With
        TBCarteira.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBCarteira.Close

txtValorTotal_vendido = Format(TotalVendido, "###,##0.00")
txtValorTotal = Format(TotalComissao, "###,##0.00")


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_old()
On Error GoTo tratar_erro

Posicao = 0
Familiatext = ""
Contador1 = 1
Lista.ListItems.Clear
Lista1.ListItems.Clear
If TBLISTA.EOF = False Then

    Posicao = TBLISTA.RecordCount
    
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Vendas_relatorios_comissao_detalhado where Codigo = " & TBLISTA!Ordem, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                With Lista.ListItems
                    .Add , , TBAbrir!CODIGO
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBAbrir!Datavendas), "", Format(TBAbrir!Datavendas, "dd/mm/yy"))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBAbrir!Ncotacao), "", TBAbrir!Ncotacao)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Revisao), "", TBAbrir!Revisao)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBAbrir!descricao_tecnica), "", TBAbrir!descricao_tecnica)
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBAbrir!quantidade), "", Format(TBAbrir!quantidade, "###,##0.0000"))
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBAbrir!preco_unitario_desconto), "", (Format(TBAbrir!preco_unitario_desconto, "###,##0.0000000000")))
                    .Item(.Count).SubItems(11) = Format(IIf(IsNull(TBAbrir!preco_lote), 0, TBAbrir!preco_lote), "###,##0.00")
                    .Item(.Count).SubItems(12) = IIf(IsNull(TBAbrir!Comissao), 0, TBAbrir!Comissao & "%")
                    .Item(.Count).SubItems(13) = IIf(IsNull(TBAbrir!ValorComissao), 0, Format(TBAbrir!ValorComissao, "###,##0.00"))
                    .Item(.Count).SubItems(14) = IIf(IsNull(TBAbrir!vend_int), "", TBAbrir!vend_int)
                    .Item(.Count).SubItems(15) = IIf(IsNull(TBAbrir!Vend_ext), "", TBAbrir!Vend_ext)
                    .Item(.Count).SubItems(16) = IIf(IsNull(TBAbrir!Liberacao), "", TBAbrir!Liberacao)
                End With
            End If
        Else
            If TBLISTA!maquina <> "" Then
                With Lista1.ListItems
                    .Add , , TBLISTA!ID
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!QtdePrev), "", Format(TBLISTA!QtdePrev, "###,##0.00"))
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!qtdeNC), "", Format(TBLISTA!qtdeNC, "###,##0.00"))
                End With
            End If
        End If
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    If optDetalhado.Value = True Then Else
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtValorTotal_vendido = Format(TBLISTA!QtdePrevista, "###,##0.00")
    txtValorTotal = Format(TBLISTA!QtdeProduzida, "###,##0.00")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
txtValorTotal_vendido = ""
txtValorTotal = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
Formulario = "Vendas/Relatórios/Comissão"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaCombo
cmbfiltrarpor.Text = "Vendedor externo"
cmbStatus.Text = "TODOS"
ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Relatórios/Comissão"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

cmbTexto.Clear
Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

    Select Case cmbfiltrarpor
        Case "Código interno": Ordenar = "Desenho"
        Case "Código de referência": Ordenar = "n_referencia"
        Case "Descrição": Ordenar = "Descricao_tecnica"
        Case "Família": Ordenar = "Familia"
        Case "Cliente": Ordenar = "Cliente"
        Case "Vendedor externo": Ordenar = "Vend_ext"
        Case "Vendedor interno": Ordenar = "Vend_int"
    End Select
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & Ordenar & " as NomeCampo1 from Vendas_relatorios_comissao_detalhado where " & Ordenar & " <> 'Null' Group by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If TBAbrir!NomeCampo1 <> "" Then cmbTexto.AddItem TBAbrir!NomeCampo1
            TBAbrir.MoveNext
        Loop
        TBAbrir.Close
    End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcDetalhado()
On Error GoTo tratar_erro

Acao = "filtrar"
If cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

Inicio = Time

Dim Mes As String

Mes = cmbdoMes.Text

        Select Case Mes
            Case "Janeiro": Mes = 1
            Case "Fevereiro": Mes = 2
            Case "Março": Mes = 3
            Case "Abril": Mes = 4
            Case "Maio": Mes = 5
            Case "Junho": Mes = 6
            Case "Julho": Mes = 7
            Case "Agosto": Mes = 8
            Case "Setembro": Mes = 9
            Case "Outubro": Mes = 10
            Case "Novembro": Mes = 11
            Case "Dezembro": Mes = 12
        End Select


If cmbStatus.Text = "FATURADO" Then
TextoFiltro1 = "(Liberacao = 'FATURADO' OR Liberacao = 'FATURADO PARCIAL') and"
TextoFiltroRel1 = "{Vendas_relatorios_comissao_detalhado.Liberacao} = ""FATURADO"" OR {Vendas_relatorios_comissao_detalhado.Liberacao} = ""FATURADO PARCIAL"" AND"
End If

If cmbStatus.Text = "VENDIDO" Then
TextoFiltro1 = "Liberacao = 'VENDIDA' and"
TextoFiltroRel1 = "{Vendas_relatorios_comissao_detalhado.Liberacao} = ""VENDIDA"" AND"
End If

If cmbStatus.Text = "TODOS" Then
TextoFiltro1 = ""
TextoFiltroRel1 = ""
End If

    Select Case cmbfiltrarpor
        Case "Código interno":
        TextoFiltro = "Desenho"
        TextoFiltroRel = "{Vendas_relatorios_comissao_detalhado.Desenho}"
        Case "Descrição":
        TextoFiltro = "Descricao_tecnica"
        TextoFiltroRel = "{Vendas_relatorios_comissao_detalhado.Descricao_tecnica}"
        Case "Cliente":
        TextoFiltro = "Cliente"
        TextoFiltroRel = "Vendas_relatorios_comissao_detalhado.Cliente}"
        Case "Vendedor externo":
        TextoFiltro = "Vend_ext"
        TextoFiltroRel = "{Vendas_relatorios_comissao_detalhado.Vend_ext}"
        Case "Vendedor interno":
        TextoFiltro = "Vend_int"
        TextoFiltroRel = "{Vendas_relatorios_comissao_detalhado.Vend_Int}"
    End Select
    
    If optPeriodo.Value = True Then
        StrSql = "Select * from Vendas_relatorios_comissao_detalhado where " & TextoFiltro1 & " " & TextoFiltro & " = '" & cmbTexto & "' and (datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by datavendas"
        FormulaRelatorio = TextoFiltroRel & " = '" & cmbTexto & "' and {Vendas_relatorios_comissao_detalhado.Datavendas} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Vendas_relatorios_comissao_detalhado.datavendas} <= Date(" & _
                            Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    
    If optMesAno.Value = True Then
        StrSql = "Select * from Vendas_relatorios_comissao_detalhado where " & TextoFiltro1 & " " & TextoFiltro & " = '" & cmbTexto & "' and Ano =  '" & cmbdoAno & "' And Mes  = '" & Mes & "' order by datavendas"
        FormulaRelatorio = TextoFiltroRel & " = '" & cmbTexto & "' AND {Vendas_relatorios_comissao_detalhado.Ano} = " & cmbdoAno.Text & " and {Vendas_relatorios_comissao_detalhado.Mes} = " & Mes & ""
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcResumido()
On Error GoTo tratar_erro


With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With

Inicio = Time

Dim Mes As String

Mes = cmbdoMes.Text

        Select Case Mes
            Case "Janeiro": Mes = 1
            Case "Fevereiro": Mes = 2
            Case "Março": Mes = 3
            Case "Abril": Mes = 4
            Case "Maio": Mes = 5
            Case "Junho": Mes = 6
            Case "Julho": Mes = 7
            Case "Agosto": Mes = 8
            Case "Setembro": Mes = 9
            Case "Outubro": Mes = 10
            Case "Novembro": Mes = 11
            Case "Dezembro": Mes = 12
        End Select


If cmbStatus.Text = "FATURADO" Then
TextoFiltro1 = "(Liberacao = 'FATURADO' OR Liberacao = 'FATURADO PARCIAL') AND"
TextoFiltroRel1 = "{Vendas_Relatorios_Historico_Resumido.Liberacao} = ""FATURADO"" OR {Vendas_Relatorios_Historico_Resumido} = ""FATURADO PARCIAL"" AND"
End If

If cmbStatus.Text = "VENDIDO" Then
TextoFiltro1 = "Liberacao = 'VENDIDA' AND"
TextoFiltroRel1 = "{Vendas_Relatorios_Historico_Resumido} = ""VENDIDA"" AND"
End If

If cmbStatus.Text = "TODOS" Then
TextoFiltro1 = ""
TextoFiltroRel1 = ""
End If


    Select Case cmbfiltrarpor
        Case "Código interno":
        TextoFiltro = "Desenho"
        TextoFiltroRel = "{Vendas_Relatorios_Historico_Resumido.Desenho}"
        Case "Descrição":
        TextoFiltro = "Descricao_tecnica"
        TextoFiltroRel = "{Vendas_Relatorios_Historico_Resumido.Descricao_tecnica}"
        Case "Cliente":
        TextoFiltro = "Cliente"
        TextoFiltroRel = "Vendas_Relatorios_Historico_Resumido.Cliente}"
        Case "Vendedor externo":
        TextoFiltro = "Vend_ext"
        TextoFiltroRel = "{Vendas_Relatorios_Historico_Resumido.Vend_ext}"
        Case "Vendedor interno":
        TextoFiltro = "Vend_int"
        TextoFiltroRel = "{Vendas_Relatorios_Historico_Resumido.Vend_Int}"
    End Select

If cmbTexto.Text <> "" Then
    If optPeriodo.Value = True Then
        StrSql = "Select * from Vendas_Relatorios_Historico_Resumido where " & TextoFiltro1 & " " & TextoFiltro & " = '" & cmbTexto & "' and (datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by datavendas"
        FormulaRelatorio = TextoFiltroRel1 & " " & TextoFiltroRel & " = '" & cmbTexto & "' and {Vendas_Relatorios_Historico_Resumido.Datavendas} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Vendas_Relatorios_Historico_Resumido.datavendas} <= Date(" & _
                            Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    
    If optMesAno.Value = True Then
        StrSql = "Select * from Vendas_Relatorios_Historico_Resumido where " & TextoFiltro1 & " " & TextoFiltro & " = '" & cmbTexto & "' and Ano =  '" & cmbdoAno & "' And Mes  = '" & Mes & "' order by datavendas"
        FormulaRelatorio = TextoFiltroRel1 & " " & TextoFiltroRel & " = '" & cmbTexto & "' AND {Vendas_Relatorios_Historico_Resumido.Ano} = " & cmbdoAno.Text & " and {Vendas_Relatorios_Historico_Resumido.Mes} = " & Mes & ""
    End If
Else
    If optPeriodo.Value = True Then
        StrSql = "Select * from Vendas_Relatorios_Historico_Resumido where " & TextoFiltro1 & " " & TextoFiltro & " <> '" & cmbTexto & "' and (datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by Datavendas, Vend_ext"
        FormulaRelatorio = TextoFiltroRel1 & " " & TextoFiltroRel & " <> '" & cmbTexto & "' and {Vendas_Relatorios_Historico_Resumido.Datavendas} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Vendas_Relatorios_Historico_Resumido.datavendas} <= Date(" & _
                            Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    
    If optMesAno.Value = True Then
        StrSql = "Select * from Vendas_Relatorios_Historico_Resumido where " & TextoFiltro1 & " " & TextoFiltro & " <> '" & cmbTexto & "' and Ano =  '" & cmbdoAno & "' And Mes  = '" & Mes & "' order by datavendas, Vend_ext"
        FormulaRelatorio = TextoFiltroRel1 & " " & TextoFiltroRel & " <> '" & cmbTexto & "' AND {Vendas_Relatorios_Historico_Resumido.Ano} = " & cmbdoAno.Text & " and {Vendas_Relatorios_Historico_Resumido.Mes} = " & Mes & ""
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

Acao = "filtrar"

If optMesAno.Value = True Then
    If cmbdoAno.Text = "" Then
        NomeCampo = "o ano para pesquisa"
        ProcVerificaAcao
        cmbdoAno.SetFocus
        Exit Sub
    End If
    If cmbdoMes.Text = "" Then
        NomeCampo = "o mês para pesquisa"
        ProcVerificaAcao
        cmbdoMes.SetFocus
        Exit Sub
    End If
    
End If

If optPeriodo.Value = True Then
    If msk_fltInicio.Value = "" Then
        NomeCampo = "a data início para pesquisa"
        ProcVerificaAcao
        msk_fltInicio.SetFocus
        Exit Sub
    End If
    If msk_fltFim.Value = "" Then
        NomeCampo = "a data final para pesquisa"
        ProcVerificaAcao
        msk_fltFim.SetFocus
        Exit Sub
    End If
    
End If

If optDetalhado.Value = True Then
    ProcDetalhado
    'Debug.print StrSql
    'Debug.print FormulaRelatorio
    ProcCarregaListaDetalhada
Else
    ProcResumido
    'Debug.print StrSql
    'Debug.print FormulaRelatorio
    ProcCarregaListaResumida
End If



intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbdoMes_Click()
On Error GoTo tratar_erro
    
            Select Case cmbdoMes.Text
                Case "Janeiro": Mes = 1
                Case "Fevereiro": Mes = 2
                Case "Março": Mes = 3
                Case "Abril": Mes = 4
                Case "Maio": Mes = 5
                Case "Junho": Mes = 6
                Case "Julho": Mes = 7
                Case "Agosto": Mes = 8
                Case "Setembro": Mes = 9
                Case "Outubro": Mes = 10
                Case "Novembro": Mes = 11
                Case "Dezembro": Mes = 12
            End Select
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If optDetalhado.Value = True Then
    Lista.Visible = True
    Lista1.Visible = False
    ProcCarregaCombo
    cmbfiltrarpor.Text = "Vendedor externo"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optMesAno_Click()
On Error GoTo tratar_erro

If optMesAno.Value = True Then
    StrSql = "Select DISTINCT Year(datavendas) as Ano from Vendas_relatorios_comissao_detalhado order by year(datavendas) Desc"
    ProcCarregaComboAno
    FPeriodo.Visible = False
    FAnoMes.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

If optPeriodo.Value = True Then
    FPeriodo.Visible = True
    FAnoMes.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If optResumido.Value = True Then
    Lista.Visible = False
    Lista1.Visible = True
    ProcCarregaCombo
    cmbfiltrarpor.Text = "Vendedor externo"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaCombo()
On Error GoTo tratar_erro

cmbfiltrarpor.Clear
cmbTexto.Clear
cmbfiltrarpor.AddItem "Vendedor externo"
cmbfiltrarpor.AddItem "Vendedor interno"
If optDetalhado.Value = True Then
    cmbfiltrarpor.AddItem "Cliente"
    cmbfiltrarpor.AddItem "Código interno"
    cmbfiltrarpor.AddItem "Código de referência"
    cmbfiltrarpor.AddItem "Descrição"
    cmbfiltrarpor.AddItem "Família"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcLocalizar
    Case 2: ProcImprimir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
