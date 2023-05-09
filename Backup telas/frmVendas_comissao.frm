VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
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
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
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
   Begin MSComctlLib.ListView Lista 
      Height          =   6960
      Left            =   60
      TabIndex        =   8
      Top             =   1935
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12277
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt. venda"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Pedido int."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Comissão"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Valor comissão"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Tag             =   "T"
         Text            =   "Vendedor int."
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Tag             =   "T"
         Text            =   "Vendedor ext."
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3175
      EndProperty
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
      TabIndex        =   20
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
         Left            =   12810
         Locked          =   -1  'True
         TabIndex        =   11
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
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Valor total da comissão."
         Top             =   375
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total comissão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12885
         TabIndex        =   22
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total vendido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9165
         TabIndex        =   21
         Top             =   180
         Width           =   1605
      End
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6960
      Left            =   60
      TabIndex        =   9
      Top             =   1935
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12277
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
         Object.Width           =   20823
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Valor vendido"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor comissão"
         Object.Width           =   2646
      EndProperty
   End
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   25
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
   Begin DrawSuite2014.USImageList USImageList1 
      Left            =   13980
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_comissao.frx":030A
      Count           =   1
   End
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   24
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
   Begin VB.Frame Frame6 
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
      Height          =   960
      Left            =   55
      TabIndex        =   13
      Top             =   960
      Width           =   1695
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
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
         TabIndex        =   1
         Top             =   555
         Width           =   1425
      End
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
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
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   1185
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
      Height          =   960
      Left            =   1770
      TabIndex        =   12
      Top             =   960
      Width           =   1455
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
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
         TabIndex        =   3
         Top             =   570
         Width           =   1155
      End
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
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
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   960
      Left            =   3240
      TabIndex        =   17
      Top             =   960
      Width           =   9885
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
         ItemData        =   "frmVendas_comissao.frx":3101
         Left            =   7860
         List            =   "frmVendas_comissao.frx":310E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Opções para filtro."
         Top             =   450
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
         ItemData        =   "frmVendas_comissao.frx":312C
         Left            =   180
         List            =   "frmVendas_comissao.frx":312E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   450
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
         ItemData        =   "frmVendas_comissao.frx":3130
         Left            =   2100
         List            =   "frmVendas_comissao.frx":3132
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
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
         Left            =   8580
         TabIndex        =   27
         Top             =   240
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
         Left            =   705
         TabIndex        =   19
         Top             =   240
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
         Left            =   4237
         TabIndex        =   18
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   13135
      TabIndex        =   14
      Top             =   960
      Width           =   2115
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   7
         ToolTipText     =   "Data final."
         Top             =   555
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
         Format          =   103546881
         CurrentDate     =   39799
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   6
         ToolTipText     =   "Data inicio."
         Top             =   195
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
         Format          =   103546881
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
         Left            =   180
         TabIndex        =   16
         Top             =   600
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
         Left            =   240
         TabIndex        =   15
         Top             =   255
         Width           =   300
      End
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11935
      TabIndex        =   23
      Top             =   8940
      Width           =   3315
   End
End
Attribute VB_Name = "frmVendas_comissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        If Lista.ListItems.Count = 0 Then Exit Sub
    Else
        If Lista1.ListItems.Count = 0 Then Exit Sub
    End If
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
End If
Vendas_Relatorio_Historico = False
Vendas_Relatorio_IndiceAtraso = False
Vendas_Relatorio_Comissao = True
Compras_Relatorio_IndiceAtraso = False
PCP_relatorios_indice_atraso = False
Manutencao_Relatorio_Historico = False
FrmMenu_impressao_padrao.Show 1

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

Sub ProcCarregaLista()
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
    contador = 0
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
                    .Item(.Count).SubItems(15) = IIf(IsNull(TBAbrir!vend_ext), "", TBAbrir!vend_ext)
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
        contador = contador + 1
        PBLista.Value = contador
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
cmbFiltrarPor.Text = "Vendedor externo"
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
If Opt_individual.Value = True Then
    Select Case cmbFiltrarPor
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
End If
Lista1.ColumnHeaders(2).Text = cmbFiltrarPor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

Acao = "filtrar"
If Opt_individual.Value = True And cmbTexto = "" Then
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
ProcLimpaCamposTotais
ProcAbrirTabelas
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")

TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by data", Conexao, adOpenKeyset, adLockOptimistic
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

Set TBCarteira = CreateObject("adodb.recordset")

If cmbStatus.Text = "FATURADO" Then
TextoFiltro1 = "(Liberacao = 'FATURADO' OR Liberacao = 'FATURADO PARCIAL') and"
End If

If cmbStatus.Text = "VENDIDO" Then
TextoFiltro1 = "Liberacao = 'VENDIDA' and"
End If

If cmbStatus.Text = "TODOS" Then
TextoFiltro1 = ""
End If

If Opt_individual.Value = True Then
    Select Case cmbFiltrarPor
        Case "Código interno": TextoFiltro = "Desenho"
        Case "": TextoFiltro = "n_referencia"
        Case "Descrição": TextoFiltro = "Descricao_tecnica"
        Case "Família": TextoFiltro = "Familia"
        Case "Cliente": TextoFiltro = "Cliente"
        Case "Vendedor externo": TextoFiltro = "Vend_ext"
        Case "Vendedor interno": TextoFiltro = "Vend_int"
    End Select
    StrSql = "Select * from Vendas_relatorios_comissao_detalhado where " & TextoFiltro1 & " " & TextoFiltro & " = '" & cmbTexto & "' and (datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by datavendas"
    Debug.Print StrSql
    
    TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
Else
    Select Case cmbFiltrarPor
        Case "Vendedor externo": TextoFiltro = "Vend_ext"
        Case "Vendedor interno": TextoFiltro = "Vend_int"
    End Select
    StrSql = "Select * from Vendas_relatorios_comissao_detalhado where " & TextoFiltro1 & " Vend_ext IS NOT NULL and Vend_int IS NOT NULL and (datavendas) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' Order by " & TextoFiltro & ", datavendas"
    Debug.Print StrSql
    
    TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
End If
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If TBCarteira.EOF = False Then
    Permitido = True
    TBCarteira.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    contador = 0
    TBCarteira.MoveFirst
    Do While TBCarteira.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If Opt_individual.Value = True And optDetalhado.Value = True Then
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            Select Case cmbFiltrarPor
                Case "Código interno": Campo = TBCarteira!Desenho
                Case "Código de referência": Campo = TBCarteira!N_referencia
                Case "Descrição": Campo = TBCarteira!descricao_tecnica
                Case "Família": Campo = TBCarteira!Familia
                Case "Cliente": Campo = TBCarteira!Cliente
                Case "Vendedor externo": Campo = TBCarteira!vend_ext
                Case "Vendedor interno": Campo = TBCarteira!vend_int
            End Select
            TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & Campo & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosResumido
        End If
        TBCarteira.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

TBProdutividade.AddNew
TBProdutividade!Ordem = TBCarteira!CODIGO
TBProdutividade!OS = TBCarteira!Cotacao
TBProdutividade!Data = TBCarteira!Datavendas
TBProdutividade!QtdePrev = IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote) 'Valor total vendido sem impostos
TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira!Comissao), 0, TBCarteira!Comissao) 'Comissão (%)
TBProdutividade!qtdeNC = IIf(IsNull(TBCarteira!ValorComissao), 0, TBCarteira!ValorComissao) 'Valor comissão
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

If TBProdutividade.EOF = True Then TBProdutividade.AddNew

If Opt_individual.Value = True Then
     Familiatext = cmbTexto
Else
    Select Case cmbFiltrarPor
        Case "Código interno":  Familiatext = TBCarteira!Desenho
        Case "Código de referência":  Familiatext = TBCarteira!N_referencia
        Case "Descrição":  Familiatext = TBCarteira!descricao_tecnica
        Case "Família":  Familiatext = TBCarteira!Familia
        Case "Cliente":  Familiatext = TBCarteira!Cliente
        Case "Vendedor externo": Familiatext = TBCarteira!vend_ext
        Case "Vendedor interno": Familiatext = TBCarteira!vend_int
    End Select
End If
TBProdutividade!maquina = Familiatext
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + IIf(IsNull(TBCarteira!preco_lote), 0, TBCarteira!preco_lote) 'Valor total vendido sem impostos
TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + IIf(IsNull(TBCarteira!ValorComissao), 0, TBCarteira!ValorComissao) 'Valor total
TBProdutividade!Data = TBCarteira!Datavendas
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

quantidade = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then TBAbrir!Texto = cmbFiltrarPor & ") : " & cmbTexto Else TBAbrir!Texto = cmbFiltrarPor
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

Qtde = 0
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select Sum(Qtdeprev) as qt, Sum(QtdeNC) as Qtde from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    TBAbrir!QtdePrevista = IIf(IsNull(TBproducao!qt), 0, TBproducao!qt) 'Valor total vendido sem impostos
    TBAbrir!QtdeProduzida = IIf(IsNull(TBproducao!Qtde), 0, TBproducao!Qtde) 'Valor total comissao
    TBproducao.MoveNext
End If
TBproducao.Close

TBAbrir.Update
TBAbrir.Close

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

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
End If

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
    cmbFiltrarPor.Text = "Vendedor externo"
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
    cmbFiltrarPor.Text = "Vendedor externo"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaCombo()
On Error GoTo tratar_erro

cmbFiltrarPor.Clear
cmbTexto.Clear
cmbFiltrarPor.AddItem "Vendedor externo"
cmbFiltrarPor.AddItem "Vendedor interno"
If optDetalhado.Value = True Then
    cmbFiltrarPor.AddItem "Cliente"
    cmbFiltrarPor.AddItem "Código interno"
    cmbFiltrarPor.AddItem "Código de referência"
    cmbFiltrarPor.AddItem "Descrição"
    cmbFiltrarPor.AddItem "Família"
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
