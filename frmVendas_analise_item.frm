VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_analise_item 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Análise crítica - Localizar produtos/serviços"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   24
      Top             =   5460
      Width           =   8805
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
         Left            =   3930
         TabIndex        =   11
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
         Left            =   2340
         TabIndex        =   10
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   6150
         TabIndex        =   15
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_analise_item.frx":0000
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   5610
         TabIndex        =   14
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_analise_item.frx":37A7
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   4500
         TabIndex        =   12
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   5070
         TabIndex        =   13
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_analise_item.frx":72BD
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   6690
         TabIndex        =   16
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmVendas_analise_item.frx":B3B7
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
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
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pág.: 0 de: 0"
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
         Left            =   7500
         TabIndex        =   27
         Top             =   240
         Width           =   945
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
         TabIndex        =   26
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
         Left            =   1650
         TabIndex        =   25
         Top             =   240
         Width           =   2190
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   55
      TabIndex        =   23
      Top             =   990
      Width           =   8805
      Begin VB.CheckBox Chk_cliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar produtos/serviços do cliente"
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
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   180
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   17
      Top             =   1470
      Width           =   8805
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   28
         Top             =   210
         WhatsThisHelpID =   210
         Width           =   4785
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.ComboBox Cmb_ordenar 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmVendas_analise_item.frx":EC83
         Left            =   6330
         List            =   "frmVendas_analise_item.frx":EC8D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Ordenar por."
         Top             =   1050
         Width           =   2265
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmVendas_analise_item.frx":ECAC
         Left            =   180
         List            =   "frmVendas_analise_item.frx":ECD4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
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
         MaxLength       =   255
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   6135
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por"
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
         Left            =   6945
         TabIndex        =   22
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
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
         Left            =   2512
         TabIndex        =   19
         Top             =   840
         Width           =   1470
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
         Left            =   1537
         TabIndex        =   18
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   21
      Top             =   0
      Width           =   9305
      _ExtentX        =   16404
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
         Name            =   "Tahoma"
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
      ButtonUseMaskColor2=   0   'False
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
         Left            =   2580
         Top             =   90
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_analise_item.frx":ED79
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   6090
      Width           =   8805
      _ExtentX        =   15531
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   2445
      Left            =   60
      TabIndex        =   4
      Top             =   3000
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   4313
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
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9375
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Un. est."
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un. com."
         Object.Width           =   1499
      EndProperty
   End
End
Attribute VB_Name = "frmVendas_analise_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Orcamento_Localizar_Prod_Serv As String 'OK

Private Sub Chk_cliente_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            If Sit_REG = 1 Then
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", False
            Else
                ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", False
            End If
        ElseIf cmbfiltrarpor = "Cliente" Then
                Set TBClientes = CreateObject("adodb.recordset")
                TBClientes.Open "Select IDCliente, NomeRazao from Clientes where NomeRazao <> 'Null' order by NomeRazao", Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    Do While TBClientes.EOF = False
                        .AddItem TBClientes!NomeRazao
                        .ItemData(.NewIndex) = TBClientes!IDCliente
                        TBClientes.MoveNext
                    Loop
                    If frmVendas_analise.txtCliente <> "" Then .Text = frmVendas_analise.txtCliente
                End If
                TBClientes.Close
            ElseIf cmbfiltrarpor = "Fornecedor" Then
                    Set TBClientes = CreateObject("adodb.recordset")
                    TBClientes.Open "Select IDCliente, Nome_Razao from Compras_fornecedores where Nome_Razao <> 'Null' order by Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        Do While TBClientes.EOF = False
                            .AddItem TBClientes!Nome_Razao
                            .ItemData(cmbfamilia.NewIndex) = TBClientes!IDCliente
                            TBClientes.MoveNext
                        Loop
                    End If
                    TBClientes.Close
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Select Case Sit_REG
    Case 1: TipoFiltro = "P.Vendas = 'True'"
    Case 2: TipoFiltro = "P.Desenho is not null"
    Case 3: TipoFiltro = "P.Instrumento = 'True'"
    Case 4: TipoFiltro = "(P.SubTipoItem = 2 or P.SubTipoItem = 3)"
End Select
If Chk_cliente.Value = 1 And frmVendas_analise.txtIDcliente <> "" Then FiltroCliente = " and C.Idcliente = " & frmVendas_analise.txtIDcliente Else FiltroCliente = ""
CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.Unidade_com"
INNERJOINTEXTO = "Select " & CamposFiltro & " from (((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Projproduto_clientes PC ON PC.Codproduto = P.codproduto) LEFT JOIN Clientes C ON C.Idcliente = PC.Idcliente) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
If Cmb_ordenar = "Código interno" Then Ordenar = "P.desenho" Else Ordenar = "P.Descricao"
TextoFltroPadrao = TipoFiltro & FiltroCliente & " AND P.Bloqueado = 'False' AND P.Dtvalidacao IS NOT NULL group by " & CamposFiltro & " order by " & Ordenar

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Then
        If cmbfiltrarpor = "Cliente" Then TextoFiltro = "PC.IDCliente" Else TextoFiltro = "IA.IDGrupo"
        StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFltroPadrao
    ElseIf cmbfiltrarpor = "Família" Then
            StrSqlLocProdPadrao = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFltroPadrao
        ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                Select Case cmbfiltrarpor
                    Case "Comprimento": TextoFiltro = "P.Comprimento"
                    Case "Largura": TextoFiltro = "P.Largura"
                    Case "Espessura": TextoFiltro = "P.Espessura"
                End Select
                valor = txtTexto
                NovoValor = Replace(valor, ",", ".")
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFltroPadrao
            Else
                Select Case cmbfiltrarpor
                    Case "Código interno": TextoFiltro = "P.desenho"
                    Case "Código de referência": TextoFiltro = "IA.N_referencia"
                    Case "Descrição": TextoFiltro = "P.descricao"
                    Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                    Case "Dureza": TextoFiltro = "P.Dureza"
                    Case "Part number": TextoFiltro = "PFAB.Part_number"
                End Select
                StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFltroPadrao
    End If
Else
    StrSqlLocProdPadrao = INNERJOINTEXTO & " where " & TextoFltroPadrao
End If
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> 2 Then
    If TBLocalizar_produto_padrao.AbsolutePage = -3 Then
        ProcExibePagina (TBLocalizar_produto_padrao.PageCount - 1)
    Else
        TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.AbsolutePage - 2
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
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
    TBLocalizar_produto_padrao.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = 1
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLocalizar_produto_padrao.AbsolutePage <> -3 Then
    If TBLocalizar_produto_padrao.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLocalizar_produto_padrao.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLocalizar_produto_padrao.AbsolutePage = TBLocalizar_produto_padrao.PageCount
ProcExibePagina (TBLocalizar_produto_padrao.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyReturn: ListView1_DblClick
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9305, 4, True

If Sit_REG <> 1 Then
    Caption = "Outros - Análise crítica - Localizar produtos"
    Frame3.Visible = False
    Frame1.Top = 990
    With ListView1
        .Top = 2520
        .Height = 2925
    End With
End If

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, 0, "Produtos/Serviços", "O", False
If Permitido = False Then cmbfiltrarpor = "Código interno"

Cmb_ordenar = "Código interno"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

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

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

Referencia = ""
If ListView1.ListItems.Count = 0 Then Exit Sub
With frmVendas_analise
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where desenho = '" & ListView1.SelectedItem & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        If Sit_REG = 1 Then
            .txtdesenho = ListView1.SelectedItem
            .ProcCarregaDadosProduto
        ElseIf Sit_REG = 2 Then
                .txtdesenho_engenharia = ListView1.SelectedItem
                .ProcCarregaDadosProdutoEng
            ElseIf Sit_REG = 3 Then
                    .txtdesenho_qualidade = ListView1.SelectedItem
                    .ProcCarregaDadosProdutoQualidade
                Else
                    .txtCodInterno_processos_item = ListView1.SelectedItem
                    .ProcCarregaDadosProdutoProcessos
        End If
    End If
    TBItem.Close
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSqlLocProdPadrao = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSqlLocProdPadrao, Conexao, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLocalizar_produto_padrao.RecordCount - IIf(Pagina > 1, (TBLocalizar_produto_padrao.PageSize * (Pagina - 1)), 0), TBLocalizar_produto_padrao.PageSize)
PBLista.Value = 1
contador = 0
Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLocalizar_produto_padrao!Desenho
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLocalizar_produto_padrao!Descricao), "", TBLocalizar_produto_padrao!Descricao)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade), "", TBLocalizar_produto_padrao!Unidade)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLocalizar_produto_padrao!Unidade_com), "", TBLocalizar_produto_padrao!Unidade_com)
    End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    contador = contador + 1
    PBLista.Value = contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Pág.: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Pág.: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If


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
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

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
