VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_FiltrarCarteira 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Ordem de faturamento | Filtrar carteira"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      Height          =   945
      Left            =   3150
      TabIndex        =   6
      Top             =   2340
      Width           =   3195
      Begin VB.OptionButton Optfim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fim"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   210
         Width           =   525
      End
      Begin VB.OptionButton Optinicio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Início"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   210
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Optmeio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   900
         TabIndex        =   30
         Top             =   210
         Width           =   585
      End
      Begin VB.OptionButton optIgual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Igual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2430
         TabIndex        =   29
         Top             =   210
         Width           =   705
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   8
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   2940
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmFaturamento_FiltrarCarteira.frx":0000
         Left            =   90
         List            =   "frmFaturamento_FiltrarCarteira.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Familia."
         Top             =   480
         Width           =   2955
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   20
      Top             =   4605
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   714
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
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
      Height          =   750
      Left            =   4440
      TabIndex        =   16
      Top             =   1590
      Width           =   1905
      Begin DrawSuite2022.USOptionButton Optdescricao 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Descrição"
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
         Value           =   -1  'True
      End
      Begin DrawSuite2022.USOptionButton optespecificacao 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   450
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   450
         Caption         =   "Descrição comercial"
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
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
      Left            =   300
      TabIndex        =   14
      Top             =   3300
      Width           =   6045
      Begin DrawSuite2022.USCheckBox Chk_tem_estoque 
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         Caption         =   "Com saldo em estoque"
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
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   705
         Left            =   4380
         TabIndex        =   15
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         DibPicture      =   "frmFaturamento_FiltrarCarteira.frx":0004
         Caption         =   "Filtrar carteira"
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
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USCheckBox chkEmpenhados 
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   570
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         Caption         =   "Empenhados"
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
      Caption         =   "Período"
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
      Left            =   300
      TabIndex        =   9
      Top             =   2340
      Width           =   2835
      Begin VB.CheckBox Chk_data 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prazo final"
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
         Left            =   1470
         TabIndex        =   19
         Top             =   240
         Value           =   2  'Grayed
         Width           =   1095
      End
      Begin VB.CheckBox Chk_data 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dt. venda"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker msk_data 
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   11
         ToolTipText     =   "Data final para pesquisa."
         Top             =   540
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
         Format          =   198639617
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_data 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   12
         ToolTipText     =   "Data início para pesquisa."
         Top             =   540
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
         Format          =   198639619
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   1380
         TabIndex        =   13
         Top             =   600
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   22
      Left            =   2640
      TabIndex        =   5
      Top             =   1590
      Width           =   1785
      Begin DrawSuite2022.USOptionButton Opt_produto_filtrar 
         Height          =   255
         Left            =   270
         TabIndex        =   21
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Produtos"
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
      Begin DrawSuite2022.USOptionButton Opt_servico_filtrar 
         Height          =   255
         Left            =   270
         TabIndex        =   22
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Serviços"
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
   End
   Begin VB.Frame frame 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   26
      Left            =   300
      TabIndex        =   4
      Top             =   1590
      Width           =   2325
      Begin DrawSuite2022.USOptionButton Opt_filtrar_ped_int 
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Ped. interno"
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
         Value           =   -1  'True
      End
      Begin DrawSuite2022.USOptionButton Opt_filtrar_ped_compra 
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   450
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         Caption         =   "Ped. compra (remessa)"
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Index           =   25
      Left            =   300
      TabIndex        =   1
      Top             =   630
      Width           =   6045
      Begin VB.ComboBox Cmb_empresa_filtro 
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
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "frmFaturamento_FiltrarCarteira.frx":3654
         Left            =   180
         List            =   "frmFaturamento_FiltrarCarteira.frx":3656
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Empresa."
         Top             =   450
         Width           =   2625
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
         ItemData        =   "frmFaturamento_FiltrarCarteira.frx":3658
         Left            =   2820
         List            =   "frmFaturamento_FiltrarCarteira.frx":3674
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   3075
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Height          =   195
         Left            =   3780
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "** Empresa **"
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
         Left            =   870
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_FiltrarCarteira.frx":36ED
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_FiltrarCarteira.frx":A86D
   End
End
Attribute VB_Name = "frmFaturamento_FiltrarCarteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

If Formulario <> "Estoque/Ordem de faturamento" Then
frmFaturamento_Prod_Serv.Lista_carteira.ListItems.Clear
frmFaturamento_Prod_Serv.Lista_carteira_faturar.ListItems.Clear
Else
frmEstoque_Ordem_Faturamento.Lista_carteira.ListItems.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_filtrar_ped_compra_Click()
On Error GoTo tratar_erro

Faturamento_PI = False

If Formulario <> "Estoque/Ordem de faturamento" Then
frmFaturamento_Prod_Serv.Lista_carteira.ListItems.Clear
frmFaturamento_Prod_Serv.Lista_carteira_faturar.ListItems.Clear
Else
frmEstoque_Ordem_Faturamento.Lista_carteira.ListItems.Clear
End If

With cmbfiltrarpor
    .Clear
    If Opt_filtrar_ped_int.Value = True Then
        .AddItem "Cliente"
        .AddItem "Pedido do cliente"
        .AddItem "Pedido interno"
        .AddItem "Programa"
    Else
        .AddItem "Fornecedor"
        .AddItem "Pedido de compra"
    End If
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .Text = "Código interno"
End With


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_produto_filtrar_Click()
On Error GoTo tratar_erro

If Formulario <> "Estoque/Ordem de faturamento" Then
frmFaturamento_Prod_Serv.Lista_carteira.ListItems.Clear
frmFaturamento_Prod_Serv.Lista_carteira_faturar.ListItems.Clear
Else
frmEstoque_Ordem_Faturamento.Lista_carteira.ListItems.Clear
End If

If Opt_servico_filtrar.Value = True Then
TipoServico = True
TipoProduto = False
Chk_tem_estoque.Value = 0
chkEmpenhados.Value = F0
Chk_tem_estoque.Visible = False
chkEmpenhados.Visible = False
Else
TipoServico = False
Tipo_Produto = True
Chk_tem_estoque.Value = 1
chkEmpenhados.Value = 1
Chk_tem_estoque.Visible = True
chkEmpenhados.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

Faturamento_PI = Opt_filtrar_ped_int.Value
Faturamento_Produtos = Opt_produto_filtrar.Value
Faturamento_Comercial = optespecificacao.Value
ID_empresa = Cmb_empresa_filtro.ItemData(Cmb_empresa_filtro.ListIndex)

If Formulario = "Estoque/Ordem de faturamento" Then
frmEstoque_Ordem_Faturamento.txtIDEmpresa = ID_empresa
Else
frmFaturamento_Prod_Serv.txtIDEmpresa = ID_empresa
End If

'If Faturamento_Produtos = True Then
'Aplicacao = "P"
'Else
'Aplicacao = "S"
'End If


TextoFiltroEmpresa = "ID_empresa = " & Cmb_empresa_filtro.ItemData(Cmb_empresa_filtro.ListIndex)
If Formulario = "Estoque/Ordem de faturamento" Then
    With msk_data(1)
        If FunVerificaDataFinal(msk_data(0).Value, .Value) = False Then
            .Value = Date
            .SetFocus
            Exit Sub
        End If
    End With

    If Opt_filtrar_ped_int.Value = True Then
        TextoFiltroEmpresaRel = "{Carteira_ordem_fat.ID_empresa} = " & Cmb_empresa_filtro.ItemData(Cmb_empresa_filtro.ListIndex)
        If Opt_produto_filtrar.Value = True Then TipoFiltro = "P" Else TipoFiltro = "S"

'=========================================================
' Com empenho
'=========================================================
        If chkEmpenhados.Value = 1 Then
            EstoqueFiltro = " and Qtde_empenhada_est > 0"
            EstoqueFiltroRel = " and {Carteira_ordem_fat.Qtde_empenhada_est} > 0"
        End If
'=========================================================
' Com saldo no estoque
'=========================================================
        If Chk_tem_estoque.Value = 1 Then
            EstoqueFiltro = EstoqueFiltro & " and SaldoEstoque > 0"
            EstoqueFiltroRel = EstoqueFiltroRel & " and {Carteira_ordem_fat.SaldoEstoque} > 0"
        End If
        
        If chkEmpenhados.Value = 0 And Chk_tem_estoque.Value = 0 Then
            EstoqueFiltro = ""
            EstoqueFiltroRel = ""
        End If
        
        If Chk_data(0).Value = 1 Then DataTexto = "Datavendas" Else DataTexto = "prazofinal"
        If Chk_data(0).Value = 1 Or Chk_data(1).Value = 1 Then
            DataFiltro = " and " & DataTexto & " Between '" & Format(msk_data(0).Value, "Short Date") & "' And '" & Format(msk_data(1).Value, "Short Date") & "'"
            DataFiltroRel = " and {Carteira_ordem_fat." & DataTexto & "} >= Date(" & Year(msk_data(0).Value) & "," & Month(msk_data(0).Value) & "," & Day(msk_data(0).Value) & ") and {Carteira_ordem_fat." & DataTexto & "} <= Date(" & _
                                        Year(msk_data(1).Value) & "," & Month(msk_data(1).Value) & "," & Day(msk_data(1).Value) & ")"
        Else
            DataFiltro = ""
            DataFiltroRel = ""
        End If
        TextoFiltroPadrao = TextoFiltroEmpresa & " and Tipo = '" & TipoFiltro & "'" & DataFiltro & EstoqueFiltro & " order by " & DataTexto & ", Desenho"
        TextoFiltroPadraoRel = TextoFiltroEmpresaRel & " and {Carteira_ordem_fat.Tipo} = '" & TipoFiltro & "'" & DataFiltroRel & EstoqueFiltroRel

        If txtTexto <> "" Or cmbTexto <> "" Then
            If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Nome Fantasia" Then
            
                If cmbfiltrarpor = "Cliente" Then
                TextoFiltro = "Cliente"
                End If
                
                If cmbfiltrarpor = "Família" Then
                TextoFiltro = "Familia"
                End If
                
                If cmbfiltrarpor = "Nome Fantasia" Then
                TextoFiltro = "nomefantasia"
                End If
                
                Strsql_Carteira_Faturamento = "Select * from Carteira_ordem_fat where " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
                FormulaRelOF = "{Carteira_ordem_fat." & TextoFiltro & "} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência": TextoFiltro = "n_referencia"
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Descricao_tecnica"
                    Case "Pedido do cliente": TextoFiltro = "PCcliente"
                    Case "Pedido interno": TextoFiltro = "Ncotacao"
                    Case "Programa": TextoFiltro = "Programatexto"
                End Select
                Strsql_Carteira_Faturamento = "Select * from Carteira_ordem_fat where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                FormulaRelOF = "{Carteira_ordem_fat." & TextoFiltro & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
            End If
        Else
            Strsql_Carteira_Faturamento = "Select * from Carteira_ordem_fat where " & TextoFiltroPadrao
            FormulaRelOF = TextoFiltroPadraoRel
        End If
    Else
        TextoFiltroEmpresaRel = "{Carteira_ordem_fat_PC.ID_empresa} = " & Cmb_empresa_filtro.ItemData(Cmb_empresa_filtro.ListIndex)
        If Chk_data(0).Value = 1 Then DataTexto = "Data" Else DataTexto = "Prazo"
        If Chk_data(0).Value = 1 Or Chk_data(1).Value = 1 Then
            DataFiltro = "and " & DataTexto & " Between '" & Format(msk_data(0).Value, "Short Date") & "' And '" & Format(msk_data(1).Value, "Short Date") & "'"
            DataFiltroRel = "and {Carteira_ordem_fat_PC." & DataTexto & "} >= Date(" & Year(msk_data(0).Value) & "," & Month(msk_data(0).Value) & "," & Day(msk_data(0).Value) & ") and {Carteira_ordem_fat_PC." & DataTexto & "} <= Date(" & _
                                        Year(msk_data(1).Value) & "," & Month(msk_data(1).Value) & "," & Day(msk_data(1).Value) & ")"
        Else
            DataFiltro = ""
        End If
        TextoFiltroPadrao = TextoFiltroEmpresa & DataFiltro & " order by " & DataTexto & ", Desenho"
        TextoFiltroPadraoRel = TextoFiltroEmpresaRel & DataFiltroRel

        If txtTexto <> "" Or cmbTexto <> "" Then
            If cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Família" Then
                If cmbfiltrarpor = "Fornecedor" Then TextoFiltro = "Fornecedor" Else TextoFiltro = "Familia"
                Strsql_Carteira_Faturamento = "Select * from Carteira_ordem_fat_PC where " & TextoFiltroEmpresa & " and " & TextoFiltro & " = '" & cmbTexto & "' and " & TextoFiltroPadrao
                FormulaRelOF = TextoFiltroEmpresaRel & " and {Carteira_ordem_fat_PC." & TextoFiltro & "} = '" & cmbTexto & "' and " & TextoFiltroPadraoRel
            Else
                Select Case cmbfiltrarpor
                    Case "Código de referência": TextoFiltro = "n_referencia"
                    Case "Código interno": TextoFiltro = "Desenho"
                    Case "Descrição": TextoFiltro = "Descricao"
                    Case "Pedido de compra": TextoFiltro = "Pedido"
                End Select
                Strsql_Carteira_Faturamento = "Select * from Carteira_ordem_fat_PC where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                FormulaRelOF = "{Carteira_ordem_fat_PC." & TextoFiltro & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
            End If
        Else
            Strsql_Carteira_Faturamento = "Select * from Carteira_ordem_fat_PC where " & TextoFiltroPadrao
            FormulaRelOF = TextoFiltroPadraoRel
        End If
    End If
Else
    TextoFiltroVal = " and NF.DtValidacaoOF IS NOT NULL and NF.int_NotaFiscal IS NULL"
    CamposFiltro = " NF.ID, E.Empresa, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, NF.Id_Int_Cliente, NF.txt_Razao_Nome, NF.RPS"
    OrdenarFiltro = " order by NF.ID"
    Select Case cmbfiltrarpor
        Case "Destinatário": TextoFiltro = "NF.txt_Razao_Nome"
        Case "Ordem de faturamento": TextoFiltro = "NF.ID"
        Case "Código interno":
            IMFFiltro = "NFP.int_cod_produto" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
            FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where " & IMFFiltro
        Case "Código de referência":
            IMFFiltro = "NFP.N_Referencia" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
            FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where " & IMFFiltro
        Case "Descrição"
            IMFFiltro = "NFP.txt_descricao" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
            FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where " & IMFFiltro
        Case "Pedido do cliente"
            IMFFiltro = "NFP.pccliente" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
            FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_detalhes_nota NFP ON NF.ID = NFP.ID_Nota) INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where " & IMFFiltro
        Case "Pedido interno":
            FiltroPadrao = "Select " & CamposFiltro & " FROM (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_proposta_nota PN ON NF.ID = PN.ID_Nota) INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where PN.proposta" & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
    End Select
    If txtTexto <> "" Then
        If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Ordem de faturamento" Then
            Strsql_Carteira_Faturamento = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where   " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and NF." & TextoFiltroEmpresa & TextoFiltroVal & " group by " & CamposFiltro & OrdenarFiltro
        Else
            Strsql_Carteira_Faturamento = FiltroPadrao & " and   NF." & TextoFiltroEmpresa & TextoFiltroVal & " group by " & CamposFiltro & OrdenarFiltro
        End If
    Else
        Strsql_Carteira_Faturamento = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where   NF." & TextoFiltroEmpresa & TextoFiltroVal & " group by " & CamposFiltro & OrdenarFiltro
    End If
End If
StrSQL_OF = Strsql_Carteira_Faturamento

If Formulario = "Estoque/Ordem de faturamento" Then
    'frmEstoque_Ordem_Faturamento.ProcCorrigeFormPedIntCompra
    frmEstoque_Ordem_Faturamento.ProcCarregaListaCarteira (1)
Else
    frmFaturamento_Prod_Serv.ProcCorrigeFormPedIntCompra
    frmFaturamento_Prod_Serv.ProcCarregaListaCarteira (1)
End If

'Debug.print FormulaRelOF

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    If Formulario = "Estoque/Ordem de faturamento" Then
        .AddItem "Cliente"
        .AddItem "Código de referência"
        .AddItem "Código interno"
        .AddItem "Descrição"
        .AddItem "Família"
        .AddItem "Pedido do cliente"
        .AddItem "Pedido interno"
        .AddItem "Programa"
        .AddItem "Nome Fantasia"
    Else
        .AddItem "Ordem de faturamento"
        .AddItem "Destinatário"
        .AddItem "Código de referência"
        .AddItem "Código interno"
        .AddItem "Descrição"
        .AddItem "Pedido do cliente"
        .AddItem "Pedido interno"
    End If
End With
ProcCarregaComboEmpresa Cmb_empresa_filtro, False
cmbfiltrarpor = "Código interno"
msk_data(0).Value = Date
msk_data(1).Value = Date
Opt_produto_filtrar.Value = True
TipoServico = False
Tipo_Produto = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If Formulario = "Estoque/Ordem de faturamento" Then
frmEstoque_Ordem_Faturamento.Lista_carteira.ListItems.Clear
Else
frmFaturamento_Prod_Serv.Lista_carteira_faturar.ListItems.Clear
End If

If cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Nome Fantasia" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    With cmbTexto
        .Clear
        If Opt_filtrar_ped_int.Value = True Then NomeViewFiltro = "Carteira_ordem_fat" Else NomeViewFiltro = "Carteira_ordem_fat_PC"
        If cmbfiltrarpor = "Cliente" Then
            CampoFiltro = "Cliente"
        ElseIf cmbfiltrarpor = "Nome Fantasia" Then
            CampoFiltro = "nomefantasia"
        ElseIf cmbfiltrarpor = "Fornecedor" Then
                CampoFiltro = "Fornecedor"
            Else
                CampoFiltro = "Familia"
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select " & CampoFiltro & " as NomeCampo from " & NomeViewFiltro & " where " & CampoFiltro & " is not null group by " & CampoFiltro, Conexao, adOpenKeyset, adLockReadOnly
        If TBAbrir.EOF = False Then
            .AddItem ""
            Do While TBAbrir.EOF = False
                .AddItem TBAbrir!NomeCampo
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End With
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_filtrar_ped_int_Click()
On Error GoTo tratar_erro

Faturamento_PI = Opt_filtrar_ped_int.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Opt_servico_filtrar_Click()
On Error GoTo tratar_erro

If Formulario <> "Estoque/Ordem de faturamento" Then
frmFaturamento_Prod_Serv.Lista_carteira.ListItems.Clear
frmFaturamento_Prod_Serv.Lista_carteira_faturar.ListItems.Clear
Else
frmEstoque_Ordem_Faturamento.Lista_carteira.ListItems.Clear
End If

If Opt_servico_filtrar.Value = True Then
TipoServico = True
TipoProduto = False
Chk_tem_estoque.Value = False
chkEmpenhados.Value = False
Chk_tem_estoque.Visible = False
chkEmpenhados.Visible = False
Else
TipoServico = False
Tipo_Produto = True
Chk_tem_estoque.Value = True
chkEmpenhados.Value = True
Chk_tem_estoque.Visible = True
chkEmpenhados.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optespecificacao_Click()
On Error GoTo tratar_erro

Faturamento_Comercial = optespecificacao.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

