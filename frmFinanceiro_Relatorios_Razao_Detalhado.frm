VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmFinanceiro_Relatorios_Razao_Detalhado 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Financeiro | Relatórios - Razão"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15285
   ClipControls    =   0   'False
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   885
      Left            =   7320
      TabIndex        =   14
      Top             =   1020
      Width           =   1335
      Begin DrawSuite2022.USOptionButton Opt_cliente 
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   300
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USOptionButton Opt_fornecedor 
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   540
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         Caption         =   "Fornecedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
   End
   Begin FlexCell.Grid GridRazao 
      Height          =   8115
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   14314
      Appearance      =   0
      BackColorBkg    =   -2147483643
      BackColorFixed  =   14737632
      BorderColor     =   14737632
      CellBorderColor =   12632256
      CellBorderColorFixed=   -2147483648
      Cols            =   5
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      GridColor       =   14737632
      Rows            =   30
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   1050
      Width           =   4605
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmFinanceiro_Relatorios_Razao_Detalhado.frx":0000
         Left            =   210
         List            =   "frmFinanceiro_Relatorios_Razao_Detalhado.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   420
         Width           =   4245
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   855
      Left            =   8670
      TabIndex        =   7
      Top             =   1050
      Width           =   6585
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmFinanceiro_Relatorios_Razao_Detalhado.frx":0004
         Left            =   120
         List            =   "frmFinanceiro_Relatorios_Razao_Detalhado.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   420
         Width           =   6345
      End
      Begin VB.TextBox Txt_ID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Código do cliente."
         Top             =   450
         Width           =   705
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   1035
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1826
      ButtonCount     =   4
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
      ButtonStyle1    =   -1
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState1    =   -1
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
      ButtonLeft3     =   93
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
      ButtonLeft4     =   131
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2520
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFinanceiro_Relatorios_Razao_Detalhado.frx":0008
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   855
      Left            =   4620
      TabIndex        =   8
      Top             =   1050
      Width           =   2685
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         ToolTipText     =   "Data final."
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Format          =   104595457
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "Data inicio."
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Format          =   104595457
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   615
         TabIndex        =   10
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
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
         Left            =   1875
         TabIndex        =   9
         Top             =   270
         Width           =   255
      End
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
      FormHeightDT    =   10530
      FormWidthDT     =   15405
      FormScaleHeightDT=   10065
      FormScaleWidthDT=   15285
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmFinanceiro_Relatorios_Razao_Detalhado"
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
    Case vbKeyF2: ProcFiltrar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcAjustaGridRazao()
On Error GoTo tratar_erro

    GridRazao.AllowUserPaste = cellTextOnly
    GridRazao.AllowUserResizing = False
    GridRazao.ExtendLastCol = True
    GridRazao.BoldFixedCell = False
    GridRazao.DisplayDateTimeMask = True
    GridRazao.DisplayFocusRect = False
    GridRazao.SelectionMode = cellSelectionByRow

    GridRazao.DrawMode = cellOwnerDraw
    GridRazao.Cols = 7 '
    
'    GridRazao.BackColor1 = RGB(231, 235, 247)
    GridRazao.BackColor2 = RGB(242, 242, 242)
    
    GridRazao.Appearance = Flat
    GridRazao.ScrollBarStyle = Flat
    GridRazao.FixedRowColStyle = Flat
    GridRazao.Cell(0, 1).Text = "Data mov"
    GridRazao.Cell(0, 2).Text = "N° documento"
    GridRazao.Cell(0, 3).Text = "Histórico"
    GridRazao.Cell(0, 4).Text = "Débito"
    GridRazao.Cell(0, 5).Text = "Crédito"
    GridRazao.Cell(0, 6).Text = "Saldo"
        
    GridRazao.Column(1).CellType = cellTextBox 'Data mov
    GridRazao.Column(1).Alignment = cellCenterCenter
        
    GridRazao.Column(2).CellType = cellTextBox 'N° documento
    GridRazao.Column(2).Alignment = cellCenterCenter
    
    GridRazao.Column(3).CellType = cellTextBox 'Histórico
    GridRazao.Column(3).Alignment = cellLeftCenter
    
    GridRazao.Column(4).CellType = cellTextBox 'Débito
    GridRazao.Column(4).Alignment = cellRightCenter 'cellHyperLink
        
    GridRazao.Column(5).CellType = cellTextBox ' Crédito
    GridRazao.Column(5).Alignment = cellRightCenter 'cellHyperLink
    
    GridRazao.Column(5).CellType = cellTextBox ' Saldo
    GridRazao.Column(6).Alignment = cellRightCenter 'cellHyperLink
       
 
    GridRazao.Column(0).Width = 20
    GridRazao.Column(1).Width = 70 ' Data mov
    GridRazao.Column(2).Width = 80 'N° documento
    GridRazao.Column(3).Width = 345 'Histórico
    GridRazao.Column(4).Width = 70 ' Débito
    GridRazao.Column(5).Width = 70 ' Crédito
    GridRazao.Column(6).Width = 70 ' Crédito
    
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15255, 4, True
Formulario = "Financeiro/Relatórios/Razão"
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, True
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcAjustaGridRazao
ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaGridRazao()
On Error GoTo tratar_erro

 Set TBLISTA = CreateObject("adodb.recordset")
 TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
 Contador = 1
 Saldo_Atual = Saldo_Anterior
 
 If TBLISTA.EOF = False Then
 Do While TBLISTA.EOF = False
 Saldo_Atual = (Saldo_Atual + TBLISTA!Credito) - TBLISTA!Debito
    GridRazao.AddItem Format(TBLISTA!Data, "Short Date") & vbTab & _
                 TBLISTA!Documento & vbTab & _
                 TBLISTA!Historico & vbTab & _
                 Format(TBLISTA!Debito, "###,##0.00") & vbTab & _
                 Format(TBLISTA!Credito, "###,##0.00") & vbTab & _
                 Format(Saldo_Atual, "###,##0.00")
                 
Contador = Contador + 1
TBLISTA.MoveNext
Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
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

If Opt_cliente.Value = True Then
    NomeView = "Financeiro_relatorios_razao_cli"
Else
    NomeView = "Financeiro_relatorios_razao_forn"
End If

TextoFiltro = ""

If cmbTexto <> "" Then
    TextoFiltro = "and ID = '" & Txt_ID & "'"
    ProcFiltrar_Unitario
Else
    procFiltrar_todos
End If



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIfrmsND v5.0"
    Exit Sub
End Sub

Private Sub procFiltrar_todos()
On Error GoTo tratar_erro

'=================================================================================
' Saldo inicial
'=================================================================================
Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select ID,RAZAO, Sum(Credito-Debito) as Saldo from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Data <= '" & Format(msk_fltInicio, "Short Date") & "' GROUP BY ID, RAZAO ORDER BY ID"
'Debug.print StrSql

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
GridRazao.rows = 2

Do While TBAbrir.EOF = False
TextoFiltro = " And ID = '" & TBAbrir!ID & "' and Razao = '" & TBAbrir!Razao & "'"
'If TBAbrir!Saldo <> 0 Then
Saldo_Anterior = IIf(IsNull(TBAbrir!Saldo), 0, TBAbrir!Saldo)
GridRazao.AddItem vbTab & vbTab & TBAbrir!Razao & vbTab & vbTab & vbTab & Format(Saldo_Anterior, "###,##0.00")
'=================================================================================
' Movimentação da conta
'=================================================================================
  StrSql = "Select data, documento, historico, debito, credito from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (" & NomeView & ".Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' " & TextoFiltro & ""
  'Debug.print StrSql
  ProcCarregaGridRazao
'=================================================================================
GridRazao.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & Format(Saldo_Atual, "###,##0.00")
'End If
TBAbrir.MoveNext
Loop
End If
StrSql = ""

TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIfrmsND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar_Unitario()
On Error GoTo tratar_erro

'=================================================================================
' Saldo inicial
'=================================================================================
Set TBAbrir = CreateObject("adodb.recordset")

StrSql = "Select Sum(Credito-Debito) as Saldo from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and ID = " & IIf(Txt_ID.Text = "", 0, Txt_ID.Text) & " and Data <= '" & Format(msk_fltInicio, "Short Date") & "'"
'Debug.print StrSql
'StrSql = "Select id, Sum(Credito-Debito) as Saldo from Financeiro_relatorios_razao_cli where ID_empresa = 1 and Data <= '07/07/2020' group by ID"

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
Saldo_Anterior = IIf(IsNull(TBAbrir!Saldo), 0, TBAbrir!Saldo)
GridRazao.rows = 1
GridRazao.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & Format(Saldo_Anterior, "###,##0.00")

'=================================================================================
' Movimentação da conta
'=================================================================================
  StrSql = "Select data, documento, historico, debito, credito from " & NomeView & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and (" & NomeView & ".Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' " & TextoFiltro & ""
  ProcCarregaGridRazao
'=================================================================================
GridRazao.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & Format(Saldo_Atual, "###,##0.00")
End If

TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIfrmsND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

 With GridRazao.PageSetup
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Tahoma"
        .HeaderFont.size = 10
        .PrintCellBorders = True
        .PrintTitleColumns = True
        .PrintFixedColumn = True
        .PrintFixedRow = True
        .PrintGridlines = True
If cmbTexto <> "" Then
        .Header = "Relatório Razão | " & cmbTexto.Text & vbNewLine & "Período de " & Format(msk_fltInicio, "Short Date") & " até " & Format(msk_fltFim, "Short Date") & vbNewLine & "Saldo inicial : " & Format(Saldo_Anterior, "###,##0.00") & " | Saldo final : " & Format(Saldo_Atual, "###,##0.00")
Else
        .Header = "Relatório Razão | Todos " & IIf(Opt_cliente.Value = True, " clientes", " fornecedores") & vbNewLine & "Período de " & Format(msk_fltInicio, "Short Date") & " até " & Format(msk_fltFim, "Short Date")
End If
        .HeaderLineStyle = cellNone
        .PaperSize = cellPaperA4
        .LeftMargin = 0.8
        .TopMargin = 3
        .RightMargin = 0.8
        .BottomMargin = 1
        .HeaderMargin = 1
        .FooterMargin = 1
        .Footer = "Pag &P de &N"
        .FooterAlignment = cellRight
        .FooterFont.Name = "Tahoma"
        .FooterFont.size = 8
    End With
GridRazao.PrintPreview 100

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    Case 1: ProcFiltrar
    Case 2: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

