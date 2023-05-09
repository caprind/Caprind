VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmManutencao_Solicitacao_Abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manutenção - Equipamentos - Localizar solicitação"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
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
   ScaleHeight     =   4590
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   4275
      Left            =   55
      TabIndex        =   0
      Top             =   30
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   6897
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Dt. solicitação"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Requisitante"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   3528
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   1
      Top             =   4320
      Width           =   12000
      _ExtentX        =   21167
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
End
Attribute VB_Name = "frmManutencao_Solicitacao_Abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: Lista_DblClick
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

If frmManutencao_menu.Optproduto.Value = True Then
TextoFiltro = "Produto = 'True'"
End If

If frmManutencao_menu.optPredial.Value = True Then
TextoFiltro = "Predial = 'True'"
End If

If frmManutencao_menu.optPosto.Value = True Then
TextoFiltro = "Produto = 'False' and Predial='False'"
End If


TBLISTA.Open "Select * from manutencao where tipo = 'S' and " & TextoFiltro & " order by Data_Solicitacao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDMaquina), "SMP", TBLISTA!IDMaquina)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "PREDIAL", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = "Solicitação"
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Data_Solicitacao), "", Format(TBLISTA!Data_Solicitacao, "dd/mm/yy"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!requisitante), "", TBLISTA!requisitante)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!setor_requisitante), "", TBLISTA!setor_requisitante)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from manutencao where Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
Var = TBAbrir!CodSol
    frmManutencao.ProcLimpaCampos
    ProcPuxaDados
    ProcCriaCodigoManutencao
End If
TBAbrir.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

With frmManutencao
    '.txttipo.Text = "Corretiva"
    .txtRequisitante.Text = IIf(IsNull(TBAbrir!requisitante), "", TBAbrir!requisitante)
    .cmbSetor_Requisitante.Text = IIf(IsNull(TBAbrir!setor_requisitante), "", TBAbrir!setor_requisitante)
    .txtData_Solicitacao.Text = IIf(IsNull(TBAbrir!Data_Solicitacao), "__/__/____", Format(TBAbrir!Data_Solicitacao, "dd/mm/yyyy"))
    .txtAprovado.Text = IIf(IsNull(TBAbrir!Aprovado), "", TBAbrir!Aprovado)
    .txtSetor_Aprovado.Text = IIf(IsNull(TBAbrir!Setor_Aprovado), "", TBAbrir!Setor_Aprovado)
    .txtLista.Text = IIf(IsNull(TBAbrir!Lista), "", TBAbrir!Lista)
    .TxtID = TBAbrir!CODIGO
    .txtIDmaquina = IIf(IsNull(TBAbrir!IDMaquina), "", TBAbrir!IDMaquina)
    .txtDescricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    .txtdata = IIf(IsNull(TBAbrir!data), "", Format(TBAbrir!data, "dd/mm/yy"))
    .txtresponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    If TBAbrir!Controlada = True Then .chkControlada.Value = 1 Else .chkControlada.Value = 0
    .ProcHabilitarPrevCorr
    .chkeletrica = IIf(TBAbrir!Eletrica = True, 1, 0)
    .chkHidraulica = IIf(TBAbrir!Hidraulica = True, 1, 0)
    .chkMecanica = IIf(TBAbrir!Mecanica = True, 1, 0)
    .chkOutros = IIf(TBAbrir!Outros = True, 1, 0)
    .optPredial = IIf(TBAbrir!Predial = True, 1, 0)
    .Optproduto = IIf(TBAbrir!Produto = True, 1, 0)
    .cmbSetorPredial.Text = IIf(IsNull(TBAbrir!Setor_Predial), "", TBAbrir!Setor_Predial)
    
    If TBAbrir!Produto = False And TBAbrir!Predial = False Then
    .optPosto = True
    Else
    .optPosto = False
    End If
    
    
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriaCodigoManutencao()
On Error GoTo tratar_erro
Dim CodigoMan As String
'Var = "S"

CodigoMan = "MAN" & Right(Var, 10)
'Set TBLISTA = CreateObject("adodb.recordset")
'TBLISTA.Open "select * from manutencao where Tipo <> '" & Var & "'", Conexao, adOpenKeyset, adLockOptimistic
'If TBLISTA.EOF = False Then
'TBLISTA.MoveLast
'
'CodigoMan = TBLISTA!codman
'CodigoMan = Right(CodigoMan, 9)
'CodigoMan = Left(CodigoMan, 6)
'
'CodigoMan = Int(CodigoMan) + 1
'    Select Case Len(CodigoMan)
'        Case 1: CodigoMan = "00000" & CodigoMan
'        Case 2: CodigoMan = "0000" & CodigoMan
'        Case 3: CodigoMan = "000" & CodigoMan
'        Case 4: CodigoMan = "00" & CodigoMan
'        Case 5: CodigoMan = "0" & CodigoMan
'    End Select
'    Ano = Right(Year(Date), 2)
'CodigoMan = "MAN-" & CodigoMan & "/" & Right(Year(Date), 2)
'Else
'    CodigoMan = "MAN-000001" & "/" & Right(Year(Date), 2)
'End If
'TBLISTA.Close
frmManutencao.txtCodigo.Text = CodigoMan

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

