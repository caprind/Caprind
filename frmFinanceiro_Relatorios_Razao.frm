VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFinanceiro_Relatorios_Razao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Financeiro | Relatórios - Razão - Resumido"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7545
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   767
      DibPicture      =   "frmFinanceiro_Relatorios_Razao.frx":0000
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
      Icon            =   "frmFinanceiro_Relatorios_Razao.frx":7180
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   3600
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   714
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   12
      Top             =   1440
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
         ItemData        =   "frmFinanceiro_Relatorios_Razao.frx":749A
         Left            =   1080
         List            =   "frmFinanceiro_Relatorios_Razao.frx":749C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   180
         Width           =   6165
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
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
         TabIndex        =   13
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   2100
      Width           =   7425
      Begin VB.TextBox Txt_ID 
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
         ItemData        =   "frmFinanceiro_Relatorios_Razao.frx":749E
         Left            =   1560
         List            =   "frmFinanceiro_Relatorios_Razao.frx":74A0
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
         BackStyle       =   0  'Transparent
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
         Left            =   3307
         TabIndex        =   14
         Top             =   180
         Width           =   1470
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   11
      Top             =   450
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2520
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFinanceiro_Relatorios_Razao.frx":74A2
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   8
      Top             =   2940
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
         Format          =   488308737
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
         Format          =   488308737
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   3690
         TabIndex        =   10
         Top             =   180
         Width           =   300
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

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
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
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
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Producao_Relatorios_Total", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!QtdePrevista = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Modulo = Formulario
TBGravar!Responsavel = pubUsuario
TBGravar!Data_inicial = msk_fltInicio
TBGravar!Data_final = msk_fltFim
TBGravar.Update

'Saldo inicial
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select ID, Razao from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (" & NomeView & ".Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' " & TextoFiltro & " group by ID, Razao", Conexao, adOpenKeyset, adLockReadOnly
If TBFIltro.EOF = False Then
    Do While TBFIltro.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!Modulo = Formulario
        TBGravar!Responsavel = pubUsuario
        TBGravar!Ordem = TBFIltro!ID
        TBGravar!maquina = TBFIltro!Razao
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Sum(Debito) as Valor from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ID = " & TBFIltro!ID & " and Razao = '" & TBFIltro!Razao & "' and Tipo = 'D' and Data < '" & Format(msk_fltInicio, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            valor = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Sum(Credito) as Valor1 from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ID = " & TBFIltro!ID & " and Razao = '" & TBFIltro!Razao & "' and Tipo = 'C' and Data < '" & Format(msk_fltInicio, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
        End If
        TBAbrir.Close
        TBGravar!QtdePrev = Format(valor - Valor1, "###,##0.00")
        TBGravar.Update
        TBGravar.Close
        TBFIltro.MoveNext
    Loop
End If
TBFIltro.Close

ProcImprimirRel "{Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "' and {Producao_Relatorios.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "'", "{" & NomeView & ".ID} = {?Pm-Producao_Relatorios.Ordem} and {" & NomeView & ".Razao} = {?Pm-Producao_Relatorios.Maquina} and {" & NomeView & ".ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " " & TextoFiltroRel & " and {" & NomeView & ".Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & NomeView & ".Data} <= Date(" & Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and {Producao_Relatorios.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

