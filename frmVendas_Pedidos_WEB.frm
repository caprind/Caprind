VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Pedidos_WEB 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo | Vendas | Pedidos | WEB"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15225
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
   Icon            =   "frmVendas_Pedidos_WEB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12845.59
   ScaleMode       =   0  'User
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtObs 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   11520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   1980
      Width           =   3705
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   30
      TabIndex        =   15
      Top             =   1020
      Width           =   15195
      Begin VB.ComboBox cmbFiltrarPor 
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
         ItemData        =   "frmVendas_Pedidos_WEB.frx":000C
         Left            =   360
         List            =   "frmVendas_Pedidos_WEB.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   450
         Width           =   4305
      End
      Begin VB.CheckBox chkFaturados 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar ""Faturados"""
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
         Height          =   240
         Left            =   9330
         TabIndex        =   17
         Top             =   480
         Width           =   1665
      End
      Begin VB.TextBox txtPedidoWeb 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5070
         TabIndex        =   16
         Top             =   450
         Width           =   3945
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   480
         Left            =   11460
         TabIndex        =   22
         Top             =   330
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   847
         Autosize        =   0   'False
         Caption         =   "Execute um duplo clique na lista para abrir o pedido WEB."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   $"frmVendas_Pedidos_WEB.frx":0049
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   6307
         TabIndex        =   20
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   27
      Left            =   0
      TabIndex        =   2
      Top             =   8850
      Width           =   15315
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
         Index           =   0
         Left            =   3780
         TabIndex        =   4
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   240
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
         Index           =   0
         Left            =   9540
         TabIndex        =   3
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Index           =   0
         Left            =   11760
         TabIndex        =   5
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_Pedidos_WEB.frx":0087
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
         Index           =   0
         Left            =   11220
         TabIndex        =   6
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_Pedidos_WEB.frx":382B
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
         Index           =   0
         Left            =   10110
         TabIndex        =   7
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
         Index           =   0
         Left            =   10680
         TabIndex        =   8
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_Pedidos_WEB.frx":7334
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
         Index           =   0
         Left            =   12300
         TabIndex        =   9
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_Pedidos_WEB.frx":B423
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   2970
         TabIndex        =   13
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   13050
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4380
         TabIndex        =   10
         Top             =   240
         Width           =   1440
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   6870
      Left            =   0
      TabIndex        =   0
      Top             =   1980
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   12118
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
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Pedido"
         Object.Width           =   1289
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Cliente"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Vendedor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Condições"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Contato"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Observações"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "CNPJ"
         Object.Width           =   0
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
      FormHeightDT    =   10200
      FormWidthDT     =   15345
      FormScaleHeightDT=   12846
      FormScaleWidthDT=   15225
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USToolBar USToolBar7 
      Height          =   975
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Gerar Pedido interno"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Gerar pedido interno "
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
      ButtonLeft2     =   46
      ButtonTop2      =   2
      ButtonWidth2    =   125
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Relatório"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Relatório (F5)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   173
      ButtonTop3      =   2
      ButtonWidth3    =   59
      ButtonHeight3   =   24
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
      ButtonLeft4     =   234
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
      ButtonLeft5     =   238
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
      ButtonLeft6     =   281
      ButtonTop6      =   2
      ButtonWidth6    =   30
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList7 
         Left            =   13980
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Pedidos_WEB.frx":ECAF
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   14
      Top             =   9480
      Width           =   15315
      _ExtentX        =   27014
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
End
Attribute VB_Name = "frmVendas_Pedidos_WEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
ProcCadOF
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
    frmVendas_Pedido_WEB.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select Codigo from Empresa where Codigo = " & IDempresa & " and SemEstoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = False Then
    If Lista.ListItems.Count > 0 Then
        ProcVerificaEstoque
    End If
End If

'=============================================================================
' Verifica se está tudo preenchido corretamente no pedido de venda
'=============================================================================
With Lista
    For InitFor = 1 To .ListItems.Count

    If .ListItems.Item(InitFor).Checked = True Then

        StrSql = "Select * from Vendas_Pedidos where ID_Pedido = " & Lista.SelectedItem
        
        FunAbreBDSite
        
        If ConexaoMySql.State = 1 Then
        
        Set TBMySQL = New ADODB.Recordset
        
        TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
         If TBMySQL.EOF = False Then
                
                If TBMySQL.Fields!Cliente = "" Then
                    USMsgBox " Erro campo nome do cliente faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!ValorTotal = "" Then
                    USMsgBox " Erro campo Valor total do pedido faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!Data = "" Then
                    USMsgBox " Erro campo Data do pedido faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!vendedor = "" Then
                    USMsgBox " Erro campo Vendedor faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!Condpagto = "" Then
                    USMsgBox " Erro campo condições de pagamento faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!contato = "" Then
                    USMsgBox " Erro campo Contato faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
        
               If TBMySQL.Fields!CNPJ = "" And TBMySQL.Fields!CPF = "" Then
                    USMsgBox " Erro CNPJ ou CPF faltando ser preenchido no pedido", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
        
        End If
        TBMySQL.Close
        
        StrSql = "Select * from Vendas_Pedido_Lista Where ID_Pedido = " & Lista.SelectedItem
        
        Set TBMySQL = New ADODB.Recordset
        
        TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
         If TBMySQL.EOF = False Then
         
         Do While TBMySQL.EOF = False
                If TBMySQL.Fields!CODIGO = "" Then
                    USMsgBox " Erro código do item código: " & TBMySQL.Fields!CODIGO & " faltando ser preenchido no pedido.", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!Unidade = "" Then
                    USMsgBox " Erro unidade do item código: " & TBMySQL.Fields!CODIGO & " faltando ser preenchido no pedido.", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!Descricao = "" Then
                    USMsgBox " Erro Descrição do item código: " & TBMySQL.Fields!CODIGO & " faltando ser preenchido no pedido.", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!vlr_unit = 0 Then
                    USMsgBox " Erro Valor unitário do item código: " & TBMySQL.Fields!CODIGO & " faltando ser preenchido no pedido.", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!qt = 0 Then
                    USMsgBox " Erro Quantidade do item código: " & TBMySQL.Fields!CODIGO & " faltando ser preenchido no pedido.", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
                
                If TBMySQL.Fields!vlr_Total = 0 Then
                    USMsgBox " Erro! Valor total do item código: " & TBMySQL.Fields!CODIGO & " faltando ser preenchido no pedido.", vbCritical, "Caprind V5.0"
                    .ListItems.Item(InitFor).Checked = False
                End If
        
        TBMySQL.MoveNext
        Loop
        
        End If
         
         End If
         TBMySQL.Close
         
End If

Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
    txtObs.Text = Lista.SelectedItem.ListSubItems.Item(8).Text
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtPedidoWeb_Change()
On Error GoTo tratar_erro

If txtPedidoWEB <> "" Then
    ProcFiltrarCarteiraWEB
Else
    Lista.ListItems.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub USToolBar7_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrarCarteiraWEB
    Case 2: ProcGerarOF
    Case 3: 'ProcImprimirCartFat
    Case 5: 'ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcVerificaEstoque()
On Error GoTo tratar_erro

'=================================================================
' Abrir BD na WEB
'=================================================================
FunAbreBDSite

If ConexaoMySql.State = 1 Then

StrSql = "SELECT * FROM Vendas_Pedido_Lista Where ID_Pedido = '" & Lista.SelectedItem & "'"
'Debug.print StrSql

Set TBLISTA = New ADODB.Recordset
'=================================================================
' Buscar produtos do pedido na WEB
'=================================================================
TBLISTA.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        Set TBEstoque = New ADODB.Recordset
        TBEstoque.Open "Select Saldo, codigo from Estoque_Controle_Saldo_Item where Saldo > 0 and Codigo = '" & TBLISTA!CODIGO & "'", Conexao, adOpenKeyset, adLockOptimistic, adCmdText
            If TBEstoque.EOF = True Then
                USMsgBox "Atenção, o item código " & TBLISTA!CODIGO & " está sem saldo no estoque."
                Exit Sub
            End If
            TBEstoque.Close
            TBLISTA.MoveNext
    Loop
End If

TBLISTA.Close
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar7 Me, 15315, 6, True
ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcGerarOF()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente cadastrar pedido(s) interno(s) dessa(s) venda(s) externa(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
        IDPedidoWEB = .ListItems.Item(InitFor).Text
        NomeRazao = .ListItems.Item(InitFor).ListSubItems(2).Text
        DataVenda = Date '.ListItems.Item(InitFor).ListSubItems(4).Text
        VendedorExterno = .ListItems.Item(InitFor).ListSubItems(5).Text
        condicoes = .ListItems.Item(InitFor).ListSubItems(6).Text
        NomeContato = .ListItems.Item(InitFor).ListSubItems(7).Text
        Observacoes = .ListItems.Item(InitFor).ListSubItems(8).Text
        CNPJCliente = .ListItems.Item(InitFor).ListSubItems(9).Text
        
        If CNPJCliente = "" Then
            USMsgBox "CNPJ | CPF não informado favor verificar.", vbCritical, "CAPRIND v5.0"
            Exit Sub
        End If
        
        If VendedorExterno = "" Then
            USMsgBox "Nome do vendedor no pedido não informado favor verificar.", vbCritical, "CAPRIND v5.0"
            Exit Sub
        End If
        
        If NomeRazao = "" Then
            USMsgBox "Nome do cliente no pedido não informado favor verificar.", vbCritical, "CAPRIND v5.0"
            Exit Sub
        End If
        
        ProcCadCliente
        ProcCadPedido
        ProcCadVendedorCliente
        'ProcCadOF
        '=================================================================
        ' Abrir BD na WEB
        '=================================================================
        FunAbreBDSite
        
        If ConexaoMySql.State = 1 Then
        
        StrSql = "SELECT * FROM Vendas_Pedidos Where ID_Pedido = '" & IDPedidoWEB & "'"
        'Debug.print StrSql
        
        Set TBMySQL = New ADODB.Recordset
        '=================================================================
        ' Marca pedido Web como faturado
        '=================================================================
        TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
        If TBMySQL.EOF = False Then
        TBMySQL!status = "Faturado"
        TBMySQL.Update
        TBMySQL.Close
        End If
        End If
        
        End If
    Next InitFor
End With

USMsgBox ("Novo(s) pedido(s) cadastrado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
ProcFiltrarCarteiraWEB
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadContatoCliente()
On Error GoTo tratar_erro

If NomeContato <> "" And IDCliente <> 0 Then
    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from Clientes_Contatos where IDCliente = '" & IDCliente & "' and NomeContato = '" & NomeContato & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBClientes.EOF = True Then
        TBClientes.AddNew
    End If
    
    TBClientes!IDCliente = IDCliente
    TBClientes!NomeContato = NomeContato
    TBClientes.Update
    TBClientes.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadCliente()
On Error GoTo tratar_erro

SintTelefone = ""
Email = ""
UF = ""
Bairro = ""
Endereco = ""
Numero = ""
CEP = ""
Cidade = ""
idTipoEmpresa = ""
NomeFantasia = ""
RegimeTributario = ""
RG_IE = ""

'=================================================================
' Abrir BD na WEB
'=================================================================
FunAbreBDSite

If ConexaoMySql.State = 1 Then
CNPJCliente = ReturnNumbersOnly(CNPJCliente)
CNPJ_Empresa = ReturnNumbersOnly(CNPJ_Empresa)

If Len(CNPJCliente) = 14 Then
StrSql = "SELECT * FROM Vendas_Clientes Where Cnpj = '" & CNPJCliente & "' and cnpj_Empresa = '" & Format(CNPJ_Empresa, "@@.@@@.@@@/@@@@-@@") & "'"
Else
StrSql = "SELECT * FROM Vendas_Clientes Where CPF = '" & CNPJCliente & "' and cnpj_Empresa = '" & Format(CNPJ_Empresa, "@@.@@@.@@@/@@@@-@@") & "'"
End If
'Debug.print StrSql

Set TBMySQL = New ADODB.Recordset
'=================================================================
' Buscar dados do cliente na WEB
'=================================================================
TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBMySQL.EOF = False Then

'CNPJCliente = TBMySQL!CNPJ
SintTelefone = TBMySQL!telefone
Email = TBMySQL!Email
'NomeContato = IIf(IsNull(TBMySQL!Contato1), "", TBMySQL!Contato1)
NomeRazao = IIf(IsNull(TBMySQL!Cliente), "", TBMySQL!Cliente)
NomeRazao = UCase(NomeRazao)

UF = IIf(IsNull(TBMySQL!UF), "", TBMySQL!UF)
UF = UCase(UF)
'Debug.print UF

Bairro = IIf(IsNull(TBMySQL!Bairro), "", TBMySQL!Bairro)
Bairro = UCase(Bairro)

Endereco = IIf(IsNull(TBMySQL!Endereco), "", TBMySQL!Endereco)
Endereco = UCase(Endereco)

Numero = IIf(IsNull(TBMySQL!Numero), "", TBMySQL!Numero)
CEP = IIf(IsNull(TBMySQL!CEP), "", TBMySQL!CEP)

Cidade = IIf(IsNull(TBMySQL!Cidade), "", TBMySQL!Cidade)
Cidade = RemoveAccents(Cidade)
Cidade = UCase(Mid$(Cidade, 1, 1)) & Right$(Cidade, Len(Cidade) - 1)
Cidade = Replace(Cidade, "SANTA BARBARA D'OESTE", "SANTA BARBARA DO OESTE")
Cidade = Replace(Cidade, "'", "")


TipoEmpresa = TBMySQL!Tipo

NomeContato = IIf(IsNull(TBMySQL!contato1), "", TBMySQL!contato1)
NomeContato = UCase(NomeContato)

'NomeFantasia = IIf(IsNull(TBMySQL!Cliente), "", TBMySQL!Cliente)
'RegimeTributario = IIf(IsNull(TBMySQL!Cliente), "", TBMySQL!Cliente)
'RG_IE = IIf(IsNull(TBMySQL!Cliente), "", TBMySQL!Cliente)
End If

TBMySQL.Close

'=================================================================
' Buscar dados do cliente no bd local
'=================================================================
Set TBCliente = CreateObject("adodb.recordset")

If Len(CNPJCliente) = 14 Then
    CNPJCliente = Format(CNPJCliente, "@@.@@@.@@@/@@@@-@@")
Else
    CNPJCliente = Format(CNPJCliente, "@@@.@@@.@@@-@@")
End If

StrSql = "Select * from Clientes where CPF_CNPJ = '" & CNPJCliente & "'"
'Debug.print StrSql

TBCliente.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBCliente.EOF = True Then
'=================================================================
' Se não achar cadastra
'=================================================================
        TBCliente.AddNew

    
    If Len(ReturnNumbersOnly(CNPJCliente)) = 14 Then
        ProcBuscarClienteNS (ReturnNumbersOnly(CNPJCliente))
    End If

    
        TBCliente!NomeRazao = NomeRazao
        TBCliente!UF = UF
        TBCliente!Bairro = Bairro
        TBCliente!Endereco = Endereco
        TBCliente!Numero = Numero
        TBCliente!CEP = CEP
        
        Cidade = RemoveAccents(Cidade)
        Cidade = LCase(Cidade)
        Cidade = Replace(Cidade, "`", "")
        
        TBCliente!Cidade = Cidade
        
If TipoEmpresa = "0" Then
        TBCliente!NomeFantasia = IIf(NomeFantasia = "", "", NomeFantasia)
        TBCliente!Presumido = IIf(RegimeTributario = "Lucro presumido", 1, 0)
        TBCliente!Simples = IIf(RegimeTributario = "Simples Nacional", 1, 0)
        TBCliente!Real = IIf(RegimeTributario = "Lucro Real", 1, 0)
        TBCliente!RG_IE = RG_IE
        TBCliente!Tipo = "JP"
Else
        TBCliente!Tipo = "FP"
End If

        TBCliente!CPF_CNPJ = CNPJCliente
        TBCliente!Categoria = "A"
        TBCliente!Tel01 = SintTelefone
        TBCliente!Email = Email
        TBCliente!status = "Liberado"
        TBCliente!Pais = "BRASIL"
        TBCliente!Data = Date
        TBCliente!Responsavel = pubUsuario
        TBCliente!Codigo_pais = "1058"
        TBCliente!DtValidacao = Date
        TBCliente!RespValidacao = pubUsuario
        TBCliente!Tipo_endereco = "Rua"
        TBCliente!Tipo_bairro = "Bairro"
        TBCliente!idTipoEmpresa = 1 'TipoEmpresa
        TBCliente!Enviar_NF = "True"
        TBCliente.Update
        IDCliente = TBCliente!ID
        Conexao.Execute "update Clientes set IDCliente = '" & IDCliente & "' where id = '" & IDCliente & "'"
        ProcCadEntrega
        ProcCadCobranca
        
        ProcCadContatoCliente
    Else
        IDCliente = TBCliente!ID
    End If
        TBCliente.Close
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadVendedorCliente()
On Error GoTo tratar_erro

If IDpedido <> 0 Then
StrSql = "select VV.id as IDVendedor, VP.IDcliente,VP.Cidade, cliente, Vend_ext, VE, VV.Comissao from vendas_proposta VP Inner join Vendas_Vendedores VV on VP.Vend_ext = VV.Vendedor Where Cotacao = " & IDpedido & " group by VV.id, VP.idcliente, VP.cidade, cliente, vend_ext,VE, VV.Comissao order by Vend_ext"
'Debug.print StrSql

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Vendas_Vendedores_Clientes where IDCliente = " & IDCliente & " and IDvendedor = " & TBAbrir!IDvendedor, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew

    TBClientes!IDCliente = TBAbrir!IDCliente
    TBClientes!IDvendedor = TBAbrir!IDvendedor
    TBClientes!tipocomissao = "C"
    TBClientes!Comissao = TBAbrir!Comissao
    TBClientes.Update
End If
End If
TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadEntrega()
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes_entrega where IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew
End If

TBClientes!IDCliente = IDCliente
TBClientes!Data = Date
TBClientes!Responsavel = pubUsuario
TBClientes!Tipo = "C"
TBClientes!CNPJ = CNPJCliente
TBClientes!Tipo_endereco = "Rua"
TBClientes!endereco_entrega = Endereco
TBClientes!Numero = Numero
TBClientes!Tipo_bairro = "Bairro"
TBClientes!bairro_entrega = Bairro
TBClientes!cidade_entrega = Cidade
TBClientes!uf_entrega = UF
TBClientes!cep_entrega = CEP
TBClientes.Update
TBClientes.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadCobranca()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from clientes_cobranca where idcliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
End If
TBGravar!IDCliente = IDCliente
TBGravar!Data = Date
TBGravar!Responsavel = pubUsuario
TBGravar!Tipo = "C"
TBGravar!CNPJ = CNPJCliente
TBGravar!Tipo_endereco = "Rua"
TBGravar!endereco_Cobranca = Endereco
TBGravar!Numero = Numero
TBGravar!Tipo_bairro = "Bairro"
TBGravar!bairro_Cobranca = Bairro
TBGravar!cidade_Cobranca = Cidade
TBGravar!uf_Cobranca = UF
TBGravar!cep_Cobranca = CEP
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadPedido()
On Error GoTo tratar_erro

'===============================================
' Cadastrar dados do pedido
'===============================================
ProcCadDadosPedido
'===============================================
' Cadastrar dados comerciais do pedido
'===============================================
ProcCadComPedido
'===============================================
' Cadastrar produtos do pedido
'===============================================
ProcCadProdPedido

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadDadosPedido()
On Error GoTo tratar_erro

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from Vendas_Proposta where IDPedidoWEB = '" & IDPedidoWEB & "'", Conexao, adOpenKeyset, adLockOptimistic

If TBPedido.EOF = True Then
'=================================================================
' Se não achar cadastra
'=================================================================
TBPedido.AddNew
ProcCriarNumeroProposta
TBPedido!Ncotacao = NProposta
'=================================================================
' Verifica o regimetributario da empresa emitente
'=================================================================
RegimeEmpresa_PI = FunVerifRegimeEmpresa(IDempresa)
If RegimeEmpresa_PI = 1 Then
'=================================================================
'Verifica se existe mais de uma tabela do simples cadastrada
'=================================================================
Contador = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        TabelaSN_PI = TBAbrir!Tabela
        Contador = Contador + 1
        TBAbrir.MoveNext
    Loop
    If Contador > 1 Then
        USMsgBox ("Favor informar a tabela do simples nacional utilizada para " & IIf(Vendas_Proposta = True, "essa proposta", "esse pedido") & "."), vbInformation, "CAPRIND v5.0"
        Vendas_Programacao = False
        frmVendas_proposta_tabelaSN.Show 1
    End If
End If
TBAbrir.Close

TBPedido!Tipo = "PE"
TBPedido!TabelaSN = TabelaSN_PI
TBPedido!Regime = RegimeEmpresa_PI
End If
End If

TBPedido!ID_empresa = IDempresa
TBPedido!IDPedidoWEB = IDPedidoWEB
TBPedido!status = "VENDIDA"
TBPedido!dataalteracao = Null
TBPedido!Revisao = "0"

'===============================================
' Buscar cadastro dos vendedores no bd local
'===============================================
'Debug.print VendedorExterno
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * FROM vendas_Vendedores where Vendedor like '%" & VendedorExterno & "%'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
TBPedido!vend_int = IIf(IsNull(TBUsuarios!vendedor), "", TBUsuarios!vendedor)
TBPedido!VI = IIf(IsNull(TBUsuarios!n_vendedor), "", TBUsuarios!n_vendedor)

TBPedido!Vend_ext = IIf(IsNull(TBUsuarios!vendedor), "", TBUsuarios!vendedor)
TBPedido!VE = IIf(IsNull(TBUsuarios!n_vendedor), "", TBUsuarios!n_vendedor)
Comissao = IIf(IsNull(TBUsuarios!Comissao), 0, TBUsuarios!Comissao)
End If
TBUsuarios.Close

'============================================
' Busca dados do cliente base local
'============================================
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * FROM Clientes where IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then

NomeRazao = TBClientes!NomeRazao
SintTelefone = IIf(IsNull(TBClientes!Tel01), "Não informado", TBClientes!Tel01)
SintEmail = IIf(IsNull(TBClientes!Email), "X", TBClientes!Email)

Endereco = TBClientes!Endereco
Numero = TBClientes!Numero
Bairro = TBClientes!Bairro
Cidade = TBClientes!Cidade
telefone = IIf(IsNull(TBClientes!Tel01), "Não informado", TBClientes!Tel01)
UF = TBClientes!UF
End If
TBClientes.Close


'=============================================
TBPedido!regiao = "BRASIL"
TBPedido!IDCliente = IDCliente
TBPedido!Cliente = NomeRazao
TBPedido!Remetente = ""
TBPedido!Referente = ""
TBPedido!Fax = SintTelefone
TBPedido!Email = IIf(SintEmail <> "", SintEmail, "")
TBPedido!Tipo_endereco = "RUA"
TBPedido!Endereco = IIf(Endereco <> "", Endereco, "")
TBPedido!Numero = IIf(Numero <> "", Numero, "")
TBPedido!complemento = ""
TBPedido!Tipo_bairro = "BAIRRO"
TBPedido!Bairro = IIf(Bairro <> "", Bairro, "")
TBPedido!Cidade = IIf(Cidade <> "", Cidade, "")
TBPedido!telefone = IIf(telefone <> "", telefone, "")
TBPedido!Departamento = "WEB"
TBPedido!UF = IIf(UF <> "", UCase(UF), "")
TBPedido!Tipo_cliente = "JP"
DataVenda = Format(DataVenda, "dd/mm/yyyy hh:mm:ss")
TBPedido!Datavendas = IIf(DataVenda <> "", DataVenda, "")
TBPedido!Obs = IIf(Observacoes <> "", Observacoes, "")

TBPedido!Data = Date
TBPedido!Responsavel = pubUsuario
TBPedido!DtValidacaoPI = Date
TBPedido!RespValidacaoPI = pubUsuario
TBPedido!DtValidacao = Date
TBPedido!RespValidacao = pubUsuario
TBPedido!Tipo = "PE"
TBPedido!Regime = RegimeEmpresa_PI
TBPedido!Remetente = NomeContato
TBPedido!Pedido = "WEB"

TBPedido.Update
IDpedido = TBPedido!Cotacao

Conexao.Execute "Update Vendas_proposta set ordenarproposta = " & TBPedido!Cotacao & " where cotacao = " & TBPedido!Cotacao


TBPedido.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadComPedido()
On Error GoTo tratar_erro

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * FROM vendas_comercial where cotacao = '" & IDpedido & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = True Then
TBContas.AddNew
TBContas!condicoes = condicoes
TBContas!Cotacao = IDpedido
TBContas.Update
TBContas.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadProdPedido()
On Error GoTo tratar_erro
'=================================================================
' Buscar Regiao para calculo dos impostos ICMS
'=================================================================

StrSql = "SELECT Regiao FROM Regioes Where UF = '" & UF & "'"
'Debug.print StrSql

Set TBAbrir = New ADODB.Recordset

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
If TBAbrir.EOF = False Then
    regiao = TBAbrir!regiao
Else
    regiao = "DE"
End If

TBAbrir.Close
'=================================================================
' Buscar produtos do pedido na WEB
'=================================================================

'=================================================================
' Abrir BD na WEB
'=================================================================
FunAbreBDSite


If ConexaoMySql.State = 1 Then

StrSql = "SELECT * FROM Vendas_Pedido_Lista Where ID_Pedido = '" & IDPedidoWEB & "'"
'Debug.print StrSql

Set TBMySQL = New ADODB.Recordset
'=================================================================
' Buscar produtos do pedido na WEB
'=================================================================
TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBMySQL.EOF = False Then
    Do While TBMySQL.EOF = False
    
'=================================================================
' Buscar dados do produto na bd local
'=================================================================
Set TBProduto = CreateObject("adodb.recordset")
StrSql = "SELECT PP.Desenho,PP.Classe as Familia, pp.Unidade,pp.descricao,PP.ID_CFOP1,PP.ID_CF,NCST.CST_ICMS,NCST.CST_IPI,NCST.CST_PIS,NCST.CST_Cofins FROM PROJPRODUTO PP inner join tbl_NaturezaOperacao_CST NCST on NCST.ID_CFOP = pp.ID_CFOP1 WHERE Desenho = '" & TBMySQL!CODIGO & "'"
'Debug.print StrSql

TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Set TBListaPI = CreateObject("adodb.recordset")
        TBListaPI.Open "Select * from vendas_carteira Where Desenho = '" & TBProduto!Desenho & "' AND IDPedidoWEB = '" & IDPedidoWEB & "' ", Conexao, adOpenKeyset, adLockOptimistic
        If TBListaPI.EOF = True Then
        TBListaPI.AddNew
        End If
        
        TBListaPI!Desenho = TBProduto!Desenho
        
        TBListaPI!ID_CFOP = IIf(TBProduto!ID_CFOP1 = "", 0, TBProduto!ID_CFOP1)
                
        TBListaPI!quantidade = TBMySQL!qt
        TBListaPI!Descricao = Trim(TBMySQL!Descricao)
        TBListaPI!descricao_tecnica = Trim(TBMySQL!Descricao)
        TBListaPI!Unidade = Trim(TBMySQL!Unidade)
        TBListaPI!Unidade_com = Trim(TBMySQL!Unidade)
        TBListaPI!ID_CF = IIf(TBProduto!ID_CF = "", 0, TBProduto!ID_CF)
        TBListaPI!txt_CST = IIf(TBProduto!CST_ICMS = "", 0, TBProduto!CST_ICMS)
        TBListaPI!Cotacao = IDpedido
        TBListaPI!Tipo = "P"
        TBListaPI!preco_unitario = Format(TBMySQL!vlr_unit, "###,##0.00")
        TBListaPI!preco_unitario_desconto = Format(TBMySQL!vlr_unit, "###,##0.00")
        TBListaPI!preco_lote = TBListaPI!quantidade * TBListaPI!preco_unitario 'Format(TBMySQL!vlr_total, "###,##0.00")
        
        TBListaPI!PrazoFinal = Date
        TBListaPI!Familia = TBProduto!Familia
        TBListaPI!Qtde_produzir = Format("0", "###,##0.00")
        TBListaPI!Liberacao = "VENDIDA"
        TBListaPI!IDPedidoWEB = IDPedidoWEB
        
        TBListaPI!Comissao = Comissao
        TBListaPI!ValorComissao = (TBListaPI!preco_lote * Comissao) / 100
        TBListaPI!PCCliente = "WEB"
        DataVenda = Format(DataVenda, "dd/mm/yyyy hh:mm:ss")
        TBListaPI!Datavendas = IIf(DataVenda <> "", DataVenda, "")
        TBListaPI!Obs_faturamento = IIf(Observacoes <> "", Observacoes, "")
        
'=====================================================================================
' Buscar impostos
'=====================================================================================
        Select Case regiao
            Case "DE": StrSql = "SELECT DE as pICMS, IPI as pIPI FROM NCM Where ID_CF = " & TBListaPI!ID_CF
            Case "CO": StrSql = "SELECT CO as pICMS, IPI as pIPI FROM NCM Where ID_CF = " & TBListaPI!ID_CF
            Case "NN": StrSql = "SELECT NN as pICMS, IPI as pIPI FROM NCM Where ID_CF = " & TBListaPI!ID_CF
            Case "SS": StrSql = "SELECT SS as pICMS, IPI as pIPI FROM NCM Where ID_CF = " & TBListaPI!ID_CF
        End Select
        'Debug.print StrSql
'========================================================
'Busca ICMS e IPI por região e classificação fiscal
'========================================================
        Set TBAbrir = New ADODB.Recordset
        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
        If TBAbrir.EOF = False Then
            pICMS = TBAbrir!pICMS
            pIPI = 0 'TBAbrir!pIPI
        End If
        TBAbrir.Close
'=====================================================================================
        TBListaPI!IntICMS = pICMS
        TBListaPI!int_IPI = pIPI
'=====================================================================================
' Verifica imposto do ICMS
'=====================================================================================
Select Case Right(TBListaPI!txt_CST, 2)
'00 – Tributada integralmente
Case "00":
    TBListaPI!dbl_Valor_ICMS = (TBListaPI!preco_lote * pICMS) / 100
    TBListaPI!BC_ICMS = TBListaPI!preco_lote
'10 – Tributada e com cobrança do ICMS por substituição tributária
'20 – Com redução de base de cálculo
'30 – Isenta ou não tributada e com cobrança do ICMS por substituição tributária
'40 – Isenta
'41 – Não tributada
'50 – Suspensão
'51 – Diferimento

'60 – ICMS cobrado anteriormente por substituição tributária
'70 – Com redução de base de cálculo e cobrança do ICMS por substituição tributária
'90 – Outras

End Select
'=====================================================================================
' Buscar aliquota PIS Cofins
'=====================================================================================
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select Pis_Produtos as PIS,Cofins_Produtos as Cofins from Impostos where ID_empresa = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBGravar.EOF = False Then
            pPIS = TBGravar!PIS
            pCofins = TBGravar!Cofins
        End If
        TBGravar.Close
'=====================================================================================
        TBListaPI!PIS_Prod = pPIS
        TBListaPI!Cofins_Prod = pCofins
'=====================================================================================
        TBListaPI.Update
        TBListaPI.Close

End If
    
    TBProduto.Close
    
    TBMySQL.MoveNext
Loop

End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCadOF()
On Error GoTo tratar_erro

FunAbreBDSite

If ConexaoMySql.State = 1 Then

StrSql = "SELECT * FROM Vendas_Pedidos"
'Debug.print StrSql

Set TBMySQL = New ADODB.Recordset
'=================================================================
' Buscar produtos do pedido na WEB
'=================================================================
TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
If TBMySQL.EOF = False Then
    Do While TBMySQL.EOF = False
    Set TBAbrir = New ADODB.Recordset
    StrSql = "Select Sum(Vlr_Total) as TotalPedido from Vendas_Pedido_Lista VPL Where ID_Pedido = '" & TBMySQL!ID_pedido & "'"
    TBAbrir.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
    If TBAbrir.EOF = False Then
    TBMySQL!ValorTotal = Replace(IIf(IsNull(TBAbrir!Totalpedido), 0, TBAbrir!Totalpedido), ",", ".")
    TBMySQL.Update
    End If
    TBAbrir.Close
    TBMySQL.MoveNext
    Loop
End If

TBMySQL.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcFiltrarCarteiraWEB()
On Error GoTo tratar_erro
Dim Filtro As String

Filtro = ""

Lista.ListItems.Clear

If txtPedidoWEB.Text <> "" Then
    Select Case cmbfiltrarpor.Text
        Case "Pedido": Filtro = " ID_Pedido like'" & txtPedidoWEB & "%' And "
        Case "Cliente": Filtro = " Cliente LIKE '" & txtPedidoWEB & "%' And "
        Case "Vendedor": Filtro = " Vendedor LIKE '" & txtPedidoWEB & "%' And "
        Case "CNPJ": Filtro = " CNPJ LIKE '" & ReturnNumbersOnly(txtPedidoWEB) & "%' And "
        Case "CPF": Filtro = " CPF Like'" & txtPedidoWEB & "%' And "
    End Select
End If


Set TBEmpresa = CreateObject("adodb.recordset")
TBEmpresa.Open "Select CNPJ from empresa where codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBEmpresa.EOF = False Then
CNPJ_Empresa = TBEmpresa!CNPJ
Else
Exit Sub
End If

TBEmpresa.Close


FunAbreBDSite

If ConexaoMySql.State = 1 Then
If chkFaturados.Value = False Then
    StrSql = "SELECT * FROM Vendas_Pedidos Where " & Filtro & " Status = 'Fechado' and CNPJ_Empresa = '" & CNPJ_Empresa & "' order by data desc"
Else
    StrSql = "SELECT * FROM Vendas_Pedidos Where " & Filtro & " Status = 'Faturado' and CNPJ_Empresa = '" & CNPJ_Empresa & "' order by data desc"
End If

Set TBMySQL = New ADODB.Recordset
'=================================================================
' Buscar pedidos na WEB
'=================================================================
TBMySQL.Open StrSql, ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
 If TBMySQL.EOF = False Then
 Do While TBMySQL.EOF = False
     With Lista.ListItems
         .Add , , TBMySQL.Fields!ID_pedido
         .Item(.Count).SubItems(1) = IIf(IsNull(TBMySQL.Fields!status), "", UCase(TBMySQL.Fields!status))
         .Item(.Count).SubItems(2) = IIf(IsNull(TBMySQL.Fields!Cliente), "", UCase(TBMySQL.Fields!Cliente))
         .Item(.Count).SubItems(3) = IIf(IsNull(TBMySQL.Fields!ValorTotal), "", Format(TBMySQL.Fields!ValorTotal, "###,##0.00"))
         .Item(.Count).SubItems(4) = IIf(IsNull(TBMySQL.Fields!Data), "", Format(TBMySQL.Fields!Data, "dd/mm/yy"))
         .Item(.Count).SubItems(5) = IIf(IsNull(TBMySQL.Fields!vendedor), "", UCase(TBMySQL.Fields!vendedor))
         .Item(.Count).SubItems(6) = IIf(IsNull(TBMySQL.Fields!Condpagto), "", UCase(TBMySQL.Fields!Condpagto))
         .Item(.Count).SubItems(7) = IIf(IsNull(TBMySQL.Fields!contato), "", UCase(TBMySQL.Fields!contato))
         .Item(.Count).SubItems(8) = IIf(IsNull(TBMySQL.Fields!observações), "", UCase(TBMySQL.Fields!observações))
         If TBMySQL.Fields!CNPJ = "" Or IsNull(TBMySQL.Fields!CNPJ) Then
         .Item(.Count).SubItems(9) = IIf(IsNull(TBMySQL.Fields!CPF), "", TBMySQL.Fields!CPF)
         Else
         .Item(.Count).SubItems(9) = IIf(IsNull(TBMySQL.Fields!CNPJ), "", TBMySQL.Fields!CNPJ)
         End If
         
     End With
     TBMySQL.MoveNext
 Loop
 End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
