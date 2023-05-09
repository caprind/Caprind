VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmprod_empenhos_PI 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Lista de empenhos do pedido interno"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   14610
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
   Icon            =   "frmprod_empenhos_PI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmprod_empenhos_PI.frx":000C
   ScaleHeight     =   9465
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   714
      DibPicture      =   "frmprod_empenhos_PI.frx":0316
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
      Icon            =   "frmprod_empenhos_PI.frx":65FA
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   9060
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8640
      Width           =   14325
      _ExtentX        =   25268
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   3
      Top             =   420
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   1720
      ButtonCount     =   3
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Ajuda"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Ajuda (F1)"
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
      ButtonWidth1    =   41
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Sair"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Sair (Esc)"
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
      ButtonLeft2     =   45
      ButtonTop2      =   2
      ButtonWidth2    =   30
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   5
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   24
      ButtonHeight3   =   24
      ButtonUseMaskColor3=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5730
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmprod_empenhos_PI.frx":6616
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView ListaEmpenhoVendidosRE 
      Height          =   6915
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   12197
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Lote"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Local de armazenamento"
         Object.Width           =   4877
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Corrida"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Certificado"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qtde. emp."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Qtde. saída"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView ListaEmpenhoVendidosOP 
      Height          =   6915
      Left            =   120
      TabIndex        =   2
      Top             =   1710
      Visible         =   0   'False
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   12197
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   14224
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Pr. final"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Qtde. emp."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Qtde. entr."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   2117
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   90
      TabIndex        =   4
      Top             =   1410
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RE's empenhados"
      TabPicture(0)   =   "frmprod_empenhos_PI.frx":7C41
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Ordens empenhadas"
      TabPicture(1)   =   "frmprod_empenhos_PI.frx":7C5D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
End
Attribute VB_Name = "frmprod_empenhos_PI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql_Estoque_Movimentacao_empenho As String 'OK

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
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

ProcCarregaToolBar1 Me, 14595, 3, True
ProcLimpaVariaveisPrincipais
With frmprod.listaitens.SelectedItem
    Caption = "PCP - Lista de empenhos do pedido interno (Pedido: " & .ListSubItems(18) & " - Rev.: " & .ListSubItems(19) & " - Cliente: " & .ListSubItems(20) & " - Cód. interno: " & .ListSubItems(4)
End With
SSTab1.Tab = 0
ProcCarregaListaREEmpPed

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
    
Private Sub ListaEmpenhoVendidosOP_BeforeLabelEdit(Cancel As Integer)
On Error GoTo tratar_erro

ProcOrdenaListView ListaEmpenhoVendidosOP, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoVendidosRE_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaEmpenhoVendidosRE, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    ListaEmpenhoVendidosOP.Visible = False
    With ListaEmpenhoVendidosRE
        .Visible = True
        If .Visible = True Then .SetFocus
    End With
    ProcCarregaListaREEmpPed
Else
    ListaEmpenhoVendidosRE.Visible = False
    With ListaEmpenhoVendidosOP
        .Visible = True
        .SetFocus
    End With
    ProcCarregaListaOPEmpPed
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaREEmpPed()
On Error GoTo tratar_erro

ListaEmpenhoVendidosRE.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ECEV.*, EC.* from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN Estoque_Controle EC ON ECEV.ID_Estoque = EC.IDEstoque where ECEV.ID_Carteira = '" & frmprod.listaitens.SelectedItem & "' and Qtde_empenhada - ISNULL(Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With ListaEmpenhoVendidosRE.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDEstoque), "", TBLISTA!IDEstoque)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!LOTE), "", TBLISTA!LOTE)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!local_armaz), "", TBLISTA!local_armaz)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Corrida), "", TBLISTA!Corrida)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
            valor = IIf(IsNull(TBLISTA!Qtde_empenhada), 0, TBLISTA!Qtde_empenhada)
            .Item(.Count).SubItems(7) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .Item(.Count).SubItems(8) = Format(Valor1, "###,##0.0000")
            .Item(.Count).SubItems(9) = Format(valor - Valor1, "###,##0.0000")
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaOPEmpPed()
On Error GoTo tratar_erro

ListaEmpenhoVendidosOP.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PP.*, P.* from Producao_pedidos PP INNER JOIN Producao P ON P.Ordem = PP.Ordem where PP.IDCarteira = " & frmprod.listaitens.SelectedItem & " and P.Desenho = '" & frmprod.listaitens.SelectedItem.ListSubItems(4) & "' and P.DtValidacao_custo IS NULL and PP.Qtde_empenho - ISNULL(PP.Qtde_entrada, 0) > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With ListaEmpenhoVendidosOP.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!PrazoEntrega), "", Format(TBLISTA!PrazoEntrega, "dd/mm/yy"))
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .Item(.Count).SubItems(4) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_entrada), 0, TBLISTA!Qtde_entrada)
            .Item(.Count).SubItems(5) = Valor1
            .Item(.Count).SubItems(6) = valor - Valor1
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
