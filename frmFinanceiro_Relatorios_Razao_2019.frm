VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmFinanceiro_Relatorios_Razao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Relatórios - Razão"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin FlexCell.Grid GridRazao 
      Height          =   4365
      Left            =   60
      TabIndex        =   16
      Top             =   3180
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   7699
      Cols            =   10
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   60
      TabIndex        =   12
      Top             =   990
      Width           =   7425
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
         ItemData        =   "frmFinanceiro_Relatorios_Razao.frx":0000
         Left            =   1080
         List            =   "frmFinanceiro_Relatorios_Razao.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   6165
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
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   1620
      Width           =   7425
      Begin VB.TextBox Txt_ID 
         Alignment       =   2  'Centralizar
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
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Código do cliente."
         Top             =   390
         Width           =   705
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
         ItemData        =   "frmFinanceiro_Relatorios_Razao.frx":0004
         Left            =   1560
         List            =   "frmFinanceiro_Relatorios_Razao.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   390
         Width           =   4965
      End
      Begin VB.OptionButton Opt_cliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
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
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton Opt_fornecedor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fornecedor"
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
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "ID"
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
         Left            =   6810
         TabIndex        =   15
         Top             =   180
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   3307
         TabIndex        =   14
         Top             =   180
         Width           =   1470
      End
   End
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
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
      ButtonCaption1  =   "Relatório"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Relatório (F5)"
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
      ButtonWidth1    =   51
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
      ButtonLeft2     =   55
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
      ButtonLeft3     =   59
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
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2014.USImageList USImageList1 
         Left            =   2520
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFinanceiro_Relatorios_Razao.frx":0008
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   60
      TabIndex        =   8
      Top             =   2460
      Width           =   7425
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   5940
         TabIndex        =   6
         ToolTipText     =   "Data final."
         Top             =   180
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
         Format          =   469499905
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   4080
         TabIndex        =   5
         ToolTipText     =   "Data inicio."
         Top             =   180
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
         Format          =   469499905
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Height          =   285
         Left            =   3690
         TabIndex        =   10
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
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
         Height          =   285
         Left            =   5490
         TabIndex        =   9
         Top             =   180
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmFinanceiro_Relatorios_Razao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

If cmbTexto <> "" Then Txt_ID = cmbTexto.ItemData(cmbTexto.ListIndex) Else Txt_ID = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7425, 4, True
Formulario = "Financeiro/Relatórios/Razão"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, True
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

ProcAjustaGridRazao

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAjustaGridRazao()
On Error GoTo tratar_erro

    GridRazao.AllowUserPaste = cellTextOnly
    GridRazao.AllowUserResizing = False
    GridRazao.ExtendLastCol = True
    GridRazao.BoldFixedCell = False
    GridRazao.DisplayDateTimeMask = True
    GridRazao.DisplayFocusRect = False
    GridRazao.SelectionMode = cellSelectionByRow

    GridRazao.DrawMode = cellOwnerDraw
    
    GridRazao.Appearance = Flat
    GridRazao.ScrollBarStyle = Flat
    GridRazao.FixedRowColStyle = Flat
    GridRazao.Cell(0, 1).Text = "Saldo Inicial"
    GridRazao.Cell(0, 2).Text = "Data"
    GridRazao.Cell(0, 3).Text = "Débito"
    GridRazao.Cell(0, 4).Text = "Crédito"
    GridRazao.Cell(0, 5).Text = "Saldo Final"
    GridRazao.Cell(0, 6).Text = "Documento"
    GridRazao.Cell(0, 7).Text = "Responsável"
    GridRazao.Cell(0, 8).Text = "Requisitante"
    GridRazao.Cell(0, 9).Text = "Destino"
  '  GridRazao.Cell(0, 10).Text = "PC\PI"
  '  GridRazao.Cell(0, 11).Text = "Cliente\Fornecedor"
   ' GridRazao.Cell(0, 12).Text = "Observações"
   
    GridRazao.Column(1).CellType = cellTextBox
    GridRazao.Column(1).Alignment = cellCenterCenter
        
    GridRazao.Column(2).CellType = cellDate
    GridRazao.Column(2).Alignment = cellCenterCenter
    GridRazao.Column(2).FormatString = "DD/MM/YYYY"
        
    GridRazao.Column(3).CellType = cellDate
    GridRazao.Column(3).Alignment = cellCenterCenter
    
    GridRazao.Column(4).CellType = cellTextBox
    GridRazao.Column(4).Alignment = cellRightCenter
    
    GridRazao.Column(5).CellType = cellTextBox
    GridRazao.Column(5).Alignment = cellRightCenter 'cellHyperLink
    
    GridRazao.Column(6).CellType = cellTextBox 'cellButton
    GridRazao.Column(6).Alignment = cellCenterCenter 'cellHyperLink
    
    GridRazao.Column(7).CellType = cellTextBox 'cellHyperLink
    GridRazao.Column(7).Alignment = cellCenterCenter 'cellHyperLink
    
    GridRazao.Column(8).CellType = cellTextBox 'cellHyperLink
    GridRazao.Column(8).Alignment = cellCenterCenter 'cellHyperLink
    
    GridRazao.Column(9).CellType = cellTextBox 'cellHyperLink
    GridRazao.Column(9).Alignment = cellCenterCenter 'cellHyperLink
    
'    GridRazao.Column(10).CellType = cellTextBox 'cellHyperLink
'    GridRazao.Column(10).Alignment = cellCenterCenter 'cellHyperLink
   
 
    GridRazao.Column(0).Width = 10
    GridRazao.Column(1).Width = 100
    GridRazao.Column(2).Width = 80
    GridRazao.Column(3).Width = 90
    GridRazao.Column(4).Width = 150
    GridRazao.Column(5).Width = 50
    GridRazao.Column(6).Width = 100
    GridRazao.Column(7).Width = 100
    GridRazao.Column(8).Width = 120
    GridRazao.Column(9).Width = 120
'    GridRazao.Column(10).Width = 120
   
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro
GridRazao.Rows = 1
Contador = 0

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
Acao = "visualizar impressão"
If Opt_cliente.Value = False And Opt_fornecedor.Value = False Then
    NomeCampo = "uma das opções"
    ProcVerificaAcao
    Exit Sub
End If
'ProcExcluirDadosProducaoRelatorios
'ProcExcluirDadosProducaoRelatoriosTotal
If Opt_cliente.Value = True Then
    NomeRel = "Contas_relatorio_razao_clientes.rpt"
    NomeView = "Financeiro_relatorios_razao_cli"
Else
    NomeRel = "Contas_relatorio_razao_fornecedores.rpt"
    NomeView = "Financeiro_relatorios_razao_forn"
End If
TextoFiltro = ""
TextoFiltroRel = ""
If cmbTexto <> "" Then
    TextoFiltro = "and ID = " & Txt_ID & " and Razao = '" & cmbTexto & "'"
    TextoFiltroRel = "and {" & NomeView & ".ID} = " & Txt_ID & " and {" & NomeView & ".Razao} = '" & cmbTexto & "'"
End If
'Set TBGravar = CreateObject("adodb.recordset")
'TBGravar.Open "Select * from Producao_Relatorios_Total", Conexao, adOpenKeyset, adLockOptimistic
'TBGravar.AddNew
'TBGravar!QtdePrevista = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
'TBGravar!Modulo = Formulario
'TBGravar!Responsavel = pubUsuario
'TBGravar!Data_inicial = msk_fltInicio
'TBGravar!Data_final = msk_fltFim
'TBGravar.Update

'Saldo inicial
'Set TBFIltro = CreateObject("adodb.recordset")
'TBFIltro.Open "Select ID, Razao from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (" & NomeView & ".Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' " & TextoFiltro & " group by ID, Razao", Conexao, adOpenKeyset, adLockReadOnly
'If TBFIltro.EOF = False Then
'Contador = 0
'    Do While TBFIltro.EOF = False
'
'
'
'        Set TBGravar = CreateObject("adodb.recordset")
'        TBGravar.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
'        TBGravar.AddNew
'        TBGravar!Modulo = Formulario
'        TBGravar!Responsavel = pubUsuario
'        TBGravar!Ordem = TBFIltro!ID
'        TBGravar!maquina = TBFIltro!Razao
        
        Set TBAbrir = CreateObject("adodb.recordset")
        
        TBAbrir.Open "select razao,  data, sum(Debito)+Sum(Credito) as SaldoInicial,sum(Debito) as TotalDebito, sum(Credito) as TotalCredito, sum(Debito)-Sum(Credito) as SaldoFinal from Financeiro_relatorios_razao_cli group by Data,razao order by data", Conexao, adOpenKeyset, adLockOptimistic
'        TBAbrir.Open "Select Sum(Debito) as Valor from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ID = " & TBFIltro!ID & " and Razao = '" & TBFIltro!Razao & "' and Tipo = 'D' and Data < '" & Format(msk_fltInicio, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
               Do While TBAbrir.EOF = False

'            valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
'        End If
'
'        Set TBAbrir = CreateObject("adodb.recordset")
'        TBAbrir.Open "Select Sum(Credito) as Valor1 from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ID = " & TBFIltro!ID & " and Razao = '" & TBFIltro!Razao & "' and Tipo = 'C' and Data < '" & Format(msk_fltInicio, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
'        If TBAbrir.EOF = False Then
'            Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
'        End If
'        TBAbrir.Close
'        TBGravar!QtdePrev = Format(valor - Valor1, "###,##0.00")
        
'==================================================================================

        GridRazao.AddItem SaldoInicial & vbTab & _
                 TBAbrir!data & vbTab & _
                 TBAbrir!TotalDebito & vbTab & _
                 TBAbrir!TotalCredito & vbTab & _
                 Format(TBAbrir!SaldoFinal, "###,##0.00") & vbTab & _
                 Format(TBAbrir!SaldoFinal, "###,##0.00") & vbTab
'==================================================================================
'End If

'        TBGravar.Update
'        TBGravar.Close
        TBAbrir.MoveNext
        Contador = Contador + 1
    Loop
End If
TBAbrir.Close

'ProcImprimirRel "{Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "' and {Producao_Relatorios.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "'", "{" & NomeView & ".ID} = {?Pm-Producao_Relatorios.Ordem} and {" & NomeView & ".Razao} = {?Pm-Producao_Relatorios.Maquina} and {" & NomeView & ".ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & TextoFiltroRel & " and {" & NomeView & ".Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & NomeView & ".Data} <= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and {Producao_Relatorios.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "'"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Opt_cliente_Click()
On Error GoTo tratar_erro

If Opt_cliente.Value = True Then
    With cmbTexto
        .Clear
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select IDcliente, Nome_Razao from tbl_contas_receber Group by Idcliente, Nome_Razao", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .AddItem ""
            Do While TBAbrir.EOF = False
                .AddItem TBAbrir!Nome_Razao
                .ItemData(.NewIndex) = TBAbrir!IDCliente
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        .SetFocus
    End With
    With Txt_ID
        .Text = ""
        .ToolTipText = "Código do cliente."
    End With
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Opt_fornecedor_Click()
On Error GoTo tratar_erro

If Opt_fornecedor.Value = True Then
    With cmbTexto
        .Clear
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select int_codforn, txt_Fornecedor from tbl_ContasPagar Group by int_codforn, txt_Fornecedor", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .AddItem ""
            Do While TBAbrir.EOF = False
                .AddItem TBAbrir!Txt_fornecedor
                .ItemData(.NewIndex) = TBAbrir!int_codforn
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        .SetFocus
    End With
    With Txt_ID
        .Text = ""
        .ToolTipText = "Código do fornecedor."
    End With
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

