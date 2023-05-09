VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmEstoque_consignado 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Saldo material consignado"
   ClientHeight    =   10035
   ClientLeft      =   1950
   ClientTop       =   1665
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   900
      Left            =   75
      TabIndex        =   10
      Top             =   1620
      Width           =   15165
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3060
         TabIndex        =   18
         Top             =   210
         Width           =   4785
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
            TabIndex        =   22
            Top             =   180
            Width           =   705
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
            TabIndex        =   21
            Top             =   180
            Width           =   1275
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
            TabIndex        =   20
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
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
            TabIndex        =   19
            Top             =   180
            Width           =   1155
         End
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   2775
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
         Left            =   7950
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   5145
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
         Left            =   7950
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Visible         =   0   'False
         Width           =   5145
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   1567
         TabIndex        =   12
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9787
         TabIndex        =   11
         Top             =   180
         Width           =   1470
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00E0E0E0&
      Height          =   630
      Left            =   60
      TabIndex        =   8
      Top             =   990
      Width           =   15195
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
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   13845
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Empresa :"
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
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por período"
      Height          =   195
      Left            =   11475
      TabIndex        =   4
      Top             =   1770
      Width           =   1245
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1360
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
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   9720
      Width           =   15150
      _ExtentX        =   26723
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
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   13
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
         Name            =   "Tahoma"
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
      Begin DrawSuite2014.USImageList USImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Estoque_consignado.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   11340
      TabIndex        =   14
      Top             =   1815
      Width           =   3900
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   2370
         TabIndex        =   6
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   270
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
         Format          =   132644865
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   540
         TabIndex        =   5
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   270
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
         Format          =   132644865
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1935
         TabIndex        =   16
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "De :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   300
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   7170
      Left            =   90
      TabIndex        =   17
      Top             =   2520
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   12647
      Cols            =   2
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      FixedRowColStyle=   2
      GridColor       =   12632256
      ReadOnly        =   -1  'True
      Rows            =   2
      DateFormat      =   2
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgFile 
      Height          =   240
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFolder 
      Height          =   225
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmEstoque_consignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Faturamento_Relatorios_Relacionamento    As String 'OK
Dim FiltroRel_Faturamento_Relatorios_Relacionamento    As String 'OK
Dim TBLISTA_Faturamento_Relatorios_Relacionamento      As ADODB.Recordset 'OK

'GridRelacionamento
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
'Dim CodRef As String, ValorCusto As String, DataValidacao As String, RespValidacao As String
Public IDnota As Long, IDlista As Long


Sub ProcAjuda()
On Error GoTo tratar_erro

' AbrirVideoWeb ("http://www.youtube.com/watch?v=o9mVNykTaq0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=10&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "Faturamento_relacionamento.rpt"
ProcImprimirRel FiltroRel_Faturamento_Relatorios_Relacionamento, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ID_nota, Int_codigo, int_Cod_Produto, N_Referencia, Txt_descricao, int_Qtd from tbl_Detalhes_Nota where id_nota = " & ListView1.SelectedItem & " order by Int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("Int_codigo = " & txtIDProduto)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        ProcLimparCamposProdutos
        txtIDProduto = TBLISTA!Int_codigo
        txtCodInterno = IIf(IsNull(TBLISTA!int_Cod_Produto), "", TBLISTA!int_Cod_Produto)
        txtCodref = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
        txtDescricao = IIf(IsNull(TBLISTA!txt_Descricao), "", TBLISTA!txt_Descricao)
        txtQtd = IIf(IsNull(TBLISTA!int_Qtd), "", Format(TBLISTA!int_Qtd, "###,##0.0000"))
        If ProcVerifNFComplementar(TBLISTA!ID_nota) = True Then ProcCarregaListaRelacionada True Else ProcCarregaListaRelacionada False
    Else
        USMsgBox ("Fim dos cadastros de produtos."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

'ListView1.ListItems.Clear
'ListView2.ListItems.Clear
'If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

'ListView1.ListItems.Clear
'ListView2.ListItems.Clear
'If cmbfiltrarpor = "Família produto" Or cmbfiltrarpor = "Família serviço" Then
'    txtTexto.Visible = False
'    cmbfamilia.Visible = True
'Else
    txtTexto.Visible = True
    cmbFamilia.Visible = False
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 6, True
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboFamilia cmbFamilia, "familia <> 'Null' and (compras = 'True' or vendas = 'True')", True

With cmbfiltrarpor
    .Clear
    .AddItem ""
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Emitente"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Nota fiscal"
End With



cmbfiltrarpor = "Nota fiscal"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date
'ProcCorrigeForm (False)
'Status_nota = 1
'SSTab1.Tab = 0

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaNF()
On Error GoTo tratar_erro

If StrSql_Faturamento_Relatorios_Relacionamento = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
ListView2.ListItems.Clear
Set TBLISTA_Faturamento_Relatorios_Relacionamento = CreateObject("adodb.recordset")
TBLISTA_Faturamento_Relatorios_Relacionamento.Open StrSql_Faturamento_Relatorios_Relacionamento, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Faturamento_Relatorios_Relacionamento.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListView1.ListItems.Clear
TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Faturamento_Relatorios_Relacionamento.RecordCount - IIf(Pagina > 1, (TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize * (Pagina - 1)), 0), TBLISTA_Faturamento_Relatorios_Relacionamento.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Faturamento_Relatorios_Relacionamento.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLISTA_Faturamento_Relatorios_Relacionamento!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!dt_DataEmissao), "", Format(TBLISTA_Faturamento_Relatorios_Relacionamento!dt_DataEmissao, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!int_NotaFiscal), "", TBLISTA_Faturamento_Relatorios_Relacionamento!int_NotaFiscal)
        If IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF) = False Then
            If TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
            If TBLISTA_Faturamento_Relatorios_Relacionamento!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(3) = TipoNF2
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Faturamento_Relatorios_Relacionamento!txt_Razao_Nome), "", Trim(TBLISTA_Faturamento_Relatorios_Relacionamento!txt_Razao_Nome))
    End With
    TBLISTA_Faturamento_Relatorios_Relacionamento.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Faturamento_Relatorios_Relacionamento.RecordCount
If TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
ElseIf TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount & " de: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Faturamento_Relatorios_Relacionamento.AbsolutePage - 1 & " de: " & TBLISTA_Faturamento_Relatorios_Relacionamento.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCamposProdutos()
On Error GoTo tratar_erro

txtIDProduto = ""
txtCodInterno = ""
txtCodref = ""
txtDescricao = ""
txtQtd = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaProdutos(NFcomplementar As Boolean)
On Error GoTo tratar_erro

ListView2.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Int_codigo, int_Cod_Produto, N_Referencia, Txt_descricao, int_Qtd, Saldo from tbl_Detalhes_Nota where id_nota = " & ListView1.SelectedItem & " order by int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListView2.ListItems
            .Add , , TBLISTA!Int_codigo
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!int_Cod_Produto), "", TBLISTA!int_Cod_Produto)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!txt_Descricao), "", Trim(TBLISTA!txt_Descricao))
            If NFcomplementar = False Then
                Qtde = IIf(IsNull(TBLISTA!int_Qtd), 0, TBLISTA!int_Qtd)
                .Item(.Count).SubItems(4) = Format(Qtde, "###,##0.0000")
                quantidade = IIf(IsNull(TBLISTA!Saldo), 0, TBLISTA!Saldo)
                .Item(.Count).SubItems(5) = Format(Qtde - quantidade, "###,##0.0000")
                .Item(.Count).SubItems(6) = Format(quantidade, "###,##0.0000")
            End If
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close
ProcLimparCamposProdutos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaConsignacao()
On Error GoTo tratar_erro

'Dim arrNodes2(15) As NodeData
Dim tempNode As Node
Dim intIndex, i As Integer
    
Call m_Tree.Nodes.Clear

Grid1.rows = 1

m_Row = 1
m_Col = 1
Contador1 = -1

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
'TBLISTA.Open "Select * from Estoque_Consignado_entrada order by dt_DataEmissao", Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
        Do While Not TBLISTA.EOF
          Contador1 = Contador1 + 1
          arrNodes(Contador1).Level = 0
          Saldo = TBLISTA!int_Qtd
          arrNodes(Contador1).Text = TBLISTA!int_NotaFiscal & vbTab & TBLISTA!txt_Razao_Nome & vbTab & TBLISTA!dt_DataEmissao & vbTab & TBLISTA!int_Cod_Produto & vbTab & TBLISTA!N_referencia & vbTab & TBLISTA!txt_Descricao & vbTab & Format(TBLISTA!int_Qtd, "###,##0.00") & vbTab & Format(Saldo, "###,##0.00") & vbTab & ""
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from Faturamento_Relacionamento where ID_nota_relacionada = '" & TBLISTA!ID & "' and ID_produto_relacionada = '" & TBLISTA!Int_codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
                Debug.Print TBAbrir.RecordCount
                Do While Not TBAbrir.EOF
                    If TBAbrir.EOF = False Then
                        'ProcNivel2Consignacao 'Carrega notas de saida
                    End If
                    TBAbrir.MoveNext
                Loop
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    
    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 9
        .rows = m_Row
        
        .Cell(0, 1).Text = "N° nota fiscal"
        .Cell(0, 2).Text = "Emitente"
        .Cell(0, 3).Text = "Data emissão"
        .Cell(0, 4).Text = "Cód. interno"
        .Cell(0, 5).Text = "Referencia"
        .Cell(0, 6).Text = "Descriçao"
        .Cell(0, 7).Text = "Quantidade"
        .Cell(0, 8).Text = "Saldo"
        .Column(1).Width = 120
        .Column(2).Width = 200
        .Column(3).Width = 70
        .Column(4).Width = 65
        .Column(5).Width = 60
        .Column(6).Width = 250
        .Column(7).Width = 70
        .Column(8).Width = 70
        
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellLeftCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellLeftCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellRightCenter
        
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text
        
        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            .AddItem arrNodes(intIndex).Text
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub ProcCarregaListaRelacionada(NFcomplementar As Boolean)
On Error GoTo tratar_erro

ListView3.ListItems.Clear
Qtde = IIf(txtQtd = "", 0, txtQtd)
quantidade = 0

If NFcomplementar = True Then
    TextoFiltro = "ID_nota = " & ListView1.SelectedItem & " or ID_nota_relacionada = " & ListView1.SelectedItem
Else
    TextoFiltro = "ID_nota = " & ListView1.SelectedItem & " and ID_produto = " & txtIDProduto & " or ID_nota_relacionada = " & ListView1.SelectedItem & " and ID_produto_relacionada = " & txtIDProduto
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Faturamento_Relacionamento where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBAbrir.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBAbrir.EOF = False
        With ListView3.ListItems
            .Add , , TBAbrir!ID
            
            With frmFaturamento_Prod_Serv
                If TBAbrir!ID_nota = ListView1.SelectedItem Then
                    If NFcomplementar = True Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota_relacionada Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!id_produto_relacionada
                Else
                    If NFcomplementar = True Then TextoFiltro = "NF.ID = " & TBAbrir!ID_nota Else TextoFiltro = "NFP.Int_codigo = " & TBAbrir!ID_produto
                End If
            End With
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.txt_Razao_Nome, NFP.dbl_ValorUnitario, NFP.Unidade_com from tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                .Item(.Count).SubItems(1) = IIf(IsNull(TBFI!dt_DataEmissao), "", (Format(TBFI!dt_DataEmissao, "dd/mm/yy")))
                .Item(.Count).SubItems(2) = IIf(IsNull(TBFI!int_NotaFiscal), "", TBFI!int_NotaFiscal)
                If IsNull(TBFI!TipoNF) = False Then
                    If TBFI!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
                    If TBFI!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
                    If TBFI!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
                End If
                .Item(.Count).SubItems(3) = TipoNF2
                .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!txt_Razao_Nome), "", TBFI!txt_Razao_Nome)
                
                If NFcomplementar = False Then
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!Qtde), 0, Format(TBAbrir!Qtde, "###,##0.0000"))
                    '.Item(.Count).SubItems(6) = IIf(IsNull(TBFI!dbl_ValorUnitario), 0, Format(TBFI!dbl_ValorUnitario, "###,##0.00000"))
                    '.Item(.Count).SubItems(7) = IIf(IsNull(TBFI!Unidade_com), 0, TBFI!Unidade_com)
                End If
            End If
            TBFI.Close
            
            quantidade = quantidade + TBAbrir!Qtde
        End With
        TBAbrir.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBAbrir.Close
txtQtde1 = Format(Qtde, "###,##0.0000")
txtQtdeRel = Format(quantidade, "###,##0.0000")
txtSaldo = Format(Qtde - quantidade, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

'funOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView1
    If .ListItems.Count = 0 Then Exit Sub
    If ProcVerifNFComplementar(.SelectedItem) = True Then
        With ListView2
            .ColumnHeaders(4).Width = 12235
            .ColumnHeaders(5).Width = 0
            .ColumnHeaders(6).Width = 0
            .ColumnHeaders(7).Width = 0
        End With
        ProcCarregaListaProdutos True
    Else
        With ListView2
            .ColumnHeaders(4).Width = 8635
            .ColumnHeaders(5).Width = 1200
            .ColumnHeaders(6).Width = 1200
            .ColumnHeaders(7).Width = 1200
        End With
        ProcCarregaListaProdutos False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

'OrdenaListView ListView2, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListView2
    If .ListItems.Count = 0 Then Exit Sub
    ProcLimparCamposProdutos
    txtIDProduto = .SelectedItem
    txtCodInterno = .SelectedItem.ListSubItems(1)
    txtCodref = .SelectedItem.ListSubItems(2)
    txtDescricao = .SelectedItem.ListSubItems(3)
    txtQtd = .SelectedItem.ListSubItems(4)
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

'OrdenaListView ListView3, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optEntrada_Click()
On Error GoTo tratar_erro
    
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ProcCorrigeForm (True)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCorrigeForm(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Família"
    .AddItem "Nota fiscal"
    
    If Entrada = True Then
        .AddItem "Emitente"
        .Text = "Emitente"
    Else
        .AddItem "Destinatário"
        .Text = "Destinatário"
    End If
End With

If Entrada = True Then
    With Label1
        .Caption = "Qtde. entrada"
        .Left = txtQtde1.Left + (txtQtde1.Width / 6)
    End With
    txtQtde1.ToolTipText = "Quantidade de entrada"
    ListView1.ColumnHeaders(5).Text = "Emitente"
    With ListView2
        .ColumnHeaders(5).Text = "Qtde. entr."
        .ColumnHeaders(6).Text = "Qtde. saída"
    End With
Else
    With Label1
        .Caption = "Qtde. saída"
        .Left = txtQtde1.Left + (txtQtde1.Width / 6)
    End With
    txtQtde1.ToolTipText = "Quantidade de saída"
    ListView1.ColumnHeaders(5).Text = "Destinatário"
    With ListView2
        .ColumnHeaders(5).Text = "Qtde. saída"
        .ColumnHeaders(6).Text = "Qtde. entr."
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
ListView2.ListItems.Clear
If optPeriodo.Value = 1 Then
    Frame7.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame7.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"

Select Case cmbfiltrarpor.Text
    Case "Código de referência": StrSql = "Select * from Estoque_Consignado_entrada where N_Referencia = '" & txtTexto.Text & "' order by dt_DataEmissao"
    Case "Código interno": StrSql = "Select * from Estoque_Consignado_entrada where int_Cod_Produto = '" & txtTexto.Text & "' order by dt_DataEmissao"
    Case "Emitente": StrSql = "Select * from Estoque_Consignado_entrada where txt_Razao_Nome = '" & txtTexto.Text & "' order by dt_DataEmissao"
    Case "Descrição": StrSql = "Select * from Estoque_Consignado_entrada where txt_Descricao = '" & txtTexto.Text & "' order by dt_DataEmissao"
    Case "Família"
    Case "Nota fiscal": StrSql = "Select * from Estoque_Consignado_entrada where int_NotaFiscal = '" & txtTexto.Text & "' order by dt_DataEmissao"
    Case "": StrSql = "Select * from Estoque_Consignado_entrada order by dt_DataEmissao"
End Select


ProcCarregaConsignacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

'ListView1.ListItems.Clear
'ListView2.ListItems.Clear
If txtTexto.Text <> "" And cmbfiltrarpor = "Nota fiscal" Then
    cmbFamilia.ListIndex = -1
    VerifNumero = txtTexto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto.Text = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

'If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

'Select Case ButtonIndex
'    Case 1: ProcImprimir
'    Case 2: ProcAnterior
'    Case 3: ProcProximo
'    Case 5: ProcAjuda
'    Case 6: Unload Me
'End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifNFComplementar(ID_nota As Long) As Boolean
On Error GoTo tratar_erro

ProcVerifNFComplementar = False
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & ID_nota & " and Finalidade_emissao = 2", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then ProcVerifNFComplementar = True
TBAbrir.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcCarregaComboFiltrarPor(Entrada As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    .AddItem "Emitente"
    .AddItem "Família"
    .AddItem "Nota fiscal"
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Grid1_Click()
On Error GoTo tratar_erro
Dim point As POINTAPI
Dim objCell As FlexCell.Cell
Dim intWidth As Integer

'If CheckEditStatus() Then Exit Sub
intWidth = 20

Call GetCursorPos(point)
Call ScreenToClient(Grid1.hWnd, point)
Set objCell = Grid1.HitTest(point.x, point.Y)

If Not objCell Is Nothing Then
    If objCell.Row >= m_Row And objCell.Col = m_Col Then
        Dim objNode As Node
        Set objNode = m_Tree.FindNode(objCell.Row - m_Row + 2)
        If Not objNode Is Nothing Then
            Dim i As Long, x As Long, Y As Long
            x = objCell.Left + 2 + (objNode.Level - 1) * intWidth
            Y = objCell.Top + (objCell.Height - 9) / 2
            If point.x >= x And point.x <= x + 9 And point.Y >= Y And point.Y <= Y + 9 Then
                If objNode.Expanded Then
                    objNode.Collapse
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        Grid1.RowHeight(objCell.Row + i) = 0
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                Else
                    objNode.Expand
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        If objNode.FindNode(i + 1).Visible Then
                            Grid1.RowHeight(objCell.Row + i) = -1 'DefaultRowHeight
                        End If
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                End If
            End If
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
On Error GoTo tratar_erro
Dim i As Long, j As Long
Dim x As Long, Y As Long
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim lngLevel As Long
Dim blnDrawLine As Boolean
Dim objNode As Node, tmpNode As Node
Dim intWidth As Integer
Dim intAdd As Integer

If Row < m_Row Or Col <> m_Col Then Exit Sub

intWidth = 20
intAdd = 26
    
Set objNode = m_Tree.FindNode(Row - m_Row + 2)
If Not objNode Is Nothing Then
    lngLevel = objNode.Level - 1

    'Tree lines
    hPen = CreatePen(0, 1, RGB(128, 128, 128))
    hOldPen = SelectObject(hdc, hPen)
    For i = 0 To lngLevel
        If i < lngLevel - 1 Then
            blnDrawLine = True
            Set tmpNode = objNode
            For j = i To lngLevel - 2
                Set tmpNode = tmpNode.Parent
            Next
            If tmpNode.NextNode Is Nothing Then
                blnDrawLine = False
            End If
            If blnDrawLine Then
                'All
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel - 1 Then
            'Top
            Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2)
            If Not objNode.NextNode Is Nothing Then
                'Bottom
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel Then
            'Top
            If objNode.VisibleNodesCount > 1 Then
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        End If
        'Horizontal line
        If lngLevel > 0 Then
            Call DrawLine(hdc, Left + intWidth * (lngLevel - 1) + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * (lngLevel - 1) + intAdd + 10, Top + (Bottom - Top) / 2)
        End If
    Next
    
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)

    '+/-
    If objNode.ChildrenCount > 0 Then
        hPen = CreatePen(0, 1, 0)
        hOldPen = SelectObject(hdc, hPen)
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        hOldPen = SelectObject(hdc, hBrush)
        
        x = Left + 2 + intWidth * lngLevel
        Y = Top + (Bottom - Top - 9) / 2
        
        Call Rectangle(hdc, x, Y, x + 9, Y + 9)
        If objNode.Expanded Then
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
        Else
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
            Call DrawLine(hdc, x + 4, Y + 2, x + 4, Y + 7)
        End If
    
        Call SelectObject(hdc, hOldPen)
        Call DeleteObject(hPen)
        Call SelectObject(hdc, hOldBrush)
        Call DeleteObject(hBrush)
    End If
    
    'Icon
    Debug.Print objNode.Level
    
    If objNode.Level = 1 Then
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, Image1.Picture, 16, 16, 0, 0, DI_NORMAL
    Else
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFile.Picture, 16, 16, 0, 0, DI_NORMAL 'imgFile.Picture
    End If
    
    'Text
    With Grid1.Cell(Row, Col)
        Dim rc As rect
        Call SetRect(rc, Left + intWidth * lngLevel + 37, Top, Right, Bottom)
        Call DrawText(hdc, .Text, -1, rc, DT_SINGLELINE Or DT_VCENTER)
    End With

    Handled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

