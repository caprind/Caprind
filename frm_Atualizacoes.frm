VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frm_Atualizacoes 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Atualizações - Relação de versões do sistema"
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
   Begin VB.Frame Frame9 
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
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   9360
      Width           =   15315
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
         Left            =   5880
         TabIndex        =   4
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
         Left            =   2460
         TabIndex        =   3
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin VB.ComboBox Cmb_ordenar 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_Atualizacoes.frx":0000
         Left            =   13770
         List            =   "frm_Atualizacoes.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Ordenar por."
         Top             =   210
         Width           =   1185
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   8100
         TabIndex        =   5
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Atualizacoes.frx":0029
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
         Left            =   7560
         TabIndex        =   6
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Atualizacoes.frx":37CD
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
         Left            =   6450
         TabIndex        =   7
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
         Left            =   7020
         TabIndex        =   8
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Atualizacoes.frx":72D6
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
         Left            =   8640
         TabIndex        =   9
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frm_Atualizacoes.frx":B3C5
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
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9750
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1770
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3090
         TabIndex        =   11
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12810
         TabIndex        =   10
         Top             =   240
         Width           =   885
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   14817
      Appearance      =   0
      BackColor2      =   14737632
      BackColorBkg    =   16777215
      BorderColor     =   12632256
      Cols            =   2
      DefaultFontSize =   8.25
      DisplayFocusRect=   0   'False
      ExtendLastCol   =   -1  'True
      FixedRowColStyle=   0
      GridColor       =   12632256
      ReadOnly        =   -1  'True
      Rows            =   2
      DateFormat      =   2
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   1650
      Top             =   180
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
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
      ButtonToolTipText1=   "Filtrar(F2)"
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
      ButtonLeft2     =   46
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   50
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   93
      ButtonTop4      =   2
      ButtonWidth4    =   30
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Atualizacoes.frx":EC51
         Count           =   1
      End
   End
End
Attribute VB_Name = "frm_Atualizacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Faturamento_Relatorios_Relacionamento    As String 'OK
Dim FiltroRel_Faturamento_Relatorios_Relacionamento    As String 'OK
Dim TBLISTA_Faturamento_Relatorios_Relacionamento      As ADODB.Recordset 'OK

'GridRelacionamento
Public m_Row As Long
Public m_Col As Long
Dim intIndex, i As Integer
'Dim CodRef As String, ValorCusto As String, DataValidacao As String, RespValidacao As String
Public IDlista As Long

Sub ProcAjuda()
On Error GoTo tratar_erro

' AbrirVideoWeb ("http://www.youtube.com/watch?v=o9mVNykTaq0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=10&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 3, True

If IsInternetOnline = False Then
USMsgBox "Internet indisponível no momento, por favor tente mais tarde", vbInformation, "CAPRIND v5.0"
Exit Sub
End If

FunAbreBDSite

ProcDesenhaGrid
ProcCarregaAtualizacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro
Dim L As Long

Grid1.rows = 1
TBLocalizar_produto_padrao.PageSize = IIf(txtNreg = "", 30, txtNreg)

TBLocalizar_produto_padrao.AbsolutePage = Pagina
TamanhoPagina = TBLocalizar_produto_padrao.PageSize
ContadorReg = 1

contador = 1
Contador2 = 0

Do While TBLocalizar_produto_padrao.EOF = False And (ContadorReg <= TamanhoPagina)
        With Grid1
        L = contador
            .AddItem TBLocalizar_produto_padrao!ID
            ' contador
            .Cell(L, 1).Text = IIf(IsNull(TBLocalizar_produto_padrao!versao), "", TBLocalizar_produto_padrao!versao)
            .Cell(L, 2).Text = IIf(IsNull(TBLocalizar_produto_padrao!Data), "", TBLocalizar_produto_padrao!Data)
            .Cell(L, 3).Text = IIf(IsNull(TBLocalizar_produto_padrao!Modulo), "", TBLocalizar_produto_padrao!Modulo)
            .Cell(L, 4).Text = IIf(IsNull(TBLocalizar_produto_padrao!Tipo), "", TBLocalizar_produto_padrao!Tipo)
            .Cell(L, 5).Text = IIf(IsNull(TBLocalizar_produto_padrao!Alteracoes), "", TBLocalizar_produto_padrao!Alteracoes)
            .Cell(L, 6).Text = IIf(IsNull(TBLocalizar_produto_padrao!LinkYouTube), "", TBLocalizar_produto_padrao!LinkYouTube)
            
        End With
    TBLocalizar_produto_padrao.MoveNext
    ContadorReg = ContadorReg + 1
    contador = contador + 1
    Contador2 = Contador2 + 1
Loop
lblRegistros.Caption = "Nº de registros: " & TBLocalizar_produto_padrao.RecordCount
If TBLocalizar_produto_padrao.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLocalizar_produto_padrao.PageCount
ElseIf TBLocalizar_produto_padrao.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.PageCount & " de: " & TBLocalizar_produto_padrao.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLocalizar_produto_padrao.AbsolutePage - 1 & " de: " & TBLocalizar_produto_padrao.PageCount
End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaAtualizacao()
On Error GoTo tratar_erro


Grid1.rows = 1

m_Row = 1
m_Col = 1
Contador1 = -1

If IsInternetOnline = False Then
USMsgBox "Internet indisponível no momento, por favor tente mais tarde", vbInformation, "CAPRIND v5.0"
Exit Sub
End If

FunAbreBDSite

StrSql = "Select * From Atualizacoes AT Inner Join Atualizacoes_alteracoes AA ON AA.idRevisao = AT.ID  order By Versao desc"
ProcCarregaLista


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Grid1.rows = 1
If StrSql = "" Then Exit Sub
Set TBLocalizar_produto_padrao = CreateObject("adodb.recordset")
TBLocalizar_produto_padrao.Open StrSql, ConexaoMySql, adOpenKeyset, adLockReadOnly
If TBLocalizar_produto_padrao.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesenhaGrid()
On Error GoTo tratar_erro
    
    With Grid1
        'Propriedades da grid
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .AllowUserResizing = True
        .BoldFixedCell = False
        .DisplayDateTimeMask = True
        .DisplayFocusRect = False
        .SelectionMode = cellSelectionFree
        .Cols = 7
        .rows = 1
        'Cabeçalho
        .Cell(0, 1).Text = "Versão"
        .Cell(0, 2).Text = "Data"
        .Cell(0, 3).Text = "Módulo"
        .Cell(0, 4).Text = "Tipo"
        
        .Cell(0, 5).Text = "Alteração"
        .Cell(0, 6).Text = "Documento"
        'Largura das colunas
        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 65
        .Column(3).Width = 165
        .Column(4).Width = 50
        .Column(5).Width = 600
        .Column(6).Width = 150
        'Alinhamento e tipo das celulas
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(4).CellType = cellTextBox
        .Column(5).Alignment = cellLeftCenter
        .Column(6).CellType = cellHyperLink
        .Column(6).Alignment = cellCenterCenter
        .AutoRedraw = True
        .Refresh
  End With
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_Click()
On Error GoTo tratar_erro

'USMsgBox Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col).Text

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_DblClick()
On Error GoTo tratar_erro
Dim Alteracao As String

Alteracao = Grid1.Cell(Grid1.ActiveCell.Row, 5).Text
frm_Atualizacoes_Detalhes.txtAlteracoes.Text = Alteracao
frm_Atualizacoes_Detalhes.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: 'ProcFiltrar
    'Case 2: ProcAtualizar
    'Case 4: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

