VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoque_Empenho 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque | Movimentação - Lista de empenhos"
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   14685
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
   Icon            =   "frmEstoque_Empenho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmEstoque_Empenho.frx":000C
   ScaleHeight     =   8400
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListaPedidos 
      Height          =   6855
      Left            =   210
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "ID_carteira"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1501
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   1234
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Ped. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Cliente"
         Object.Width           =   3327
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Vend. int."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Vend. ext."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Pr. final"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Qtde. vend."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qtde. emp."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde. saída"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   1587
      EndProperty
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   7995
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   767
      DibPicture      =   "frmEstoque_Empenho.frx":0316
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
      Icon            =   "frmEstoque_Empenho.frx":65FA
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   7710
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
   Begin MSComctlLib.ListView Listaordens 
      Height          =   6855
      Left            =   210
      TabIndex        =   0
      Top             =   840
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   12091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "RE"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Cód. ref."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   2632
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Cliente"
         Object.Width           =   2632
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "D"
         Text            =   "Prazo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Qtde. prod."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qtde. emp."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Qtde. saída"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Saldo"
         Object.Width           =   1587
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7500
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Empenhados por vendas"
      TabPicture(0)   =   "frmEstoque_Empenho.frx":6616
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Empenhados na produção"
      TabPicture(1)   =   "frmEstoque_Empenho.frx":6632
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
   End
End
Attribute VB_Name = "frmEstoque_Empenho"
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

'ProcCarregaToolBar1 Me, 14355, 3, True
If PCP_Ordem = True Then Caption = "PCP - Gerenciamento de ordem - Lista de empenhos"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
Listaordens.Visible = False
ListaPedidos.Visible = True
ProcCarregaListaPedidos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaordens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Listaordens, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listapedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaPedidos, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listapedidos_DblClick()
On Error GoTo tratar_erro

With ListaPedidos
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), frmestoque_item.Lista.SelectedItem.ListSubItems(4), False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then '(Vendas)
    Listaordens.Visible = False
    With ListaPedidos
        .Visible = True
    End With
    ProcCarregaListaPedidos
Else
    ListaPedidos.Visible = False '(Produção)
    With Listaordens
        .Visible = True
        If .Visible = True Then .SetFocus
    End With
    ProcCarregaListaOrdens
End If

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

Private Sub ProcCarregaListaOrdens()
On Error GoTo tratar_erro

Listaordens.ListItems.Clear
If Desenho <> "" Then TextoFiltro = "PNFC.Codinterno = '" & Desenho & "'" Else TextoFiltro = "PNFC.IDestoque = " & IDlista
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PNFC.ID, PNFC.DATA as DataEmpenho, PNFC.Responsavel as ResponsavelEmpenho, PNFC.IDestoque, PNFC.Quantidade, PNFC.Qtde_saida, P.* from Producao_NF_Consignada PNFC INNER JOIN Producao P ON PNFC.Ordem = P.Ordem where " & TextoFiltro & " and P.Status <> 'Cancelada' and P.DtValidacao_custo IS NULL and PNFC.Quantidade - ISNULL(PNFC.Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Listaordens.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!DataEmpenho), "", Format(TBLISTA!DataEmpenho, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!ResponsavelEmpenho), "", TBLISTA!ResponsavelEmpenho)
            .Item(.Count).SubItems(3) = TBLISTA!IDEstoque
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Revitem), "", TBLISTA!Revitem)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Produto), "", TBLISTA!Produto)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!PrazoEntrega), "", Format(TBLISTA!PrazoEntrega, "dd/mm/yy"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Quant), "", Format(TBLISTA!Quant, "###,##0.0000"))
            valor = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
            .Item(.Count).SubItems(12) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .Item(.Count).SubItems(13) = Format(Valor1, "###,##0.0000")
            .Item(.Count).SubItems(14) = Format(valor - Valor1, "###,##0.0000")
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

Private Sub ProcCarregaListaPedidos()
On Error GoTo tratar_erro

ListaPedidos.ListItems.Clear
If Desenho <> "" Then TextoFiltro = "VC.Desenho = '" & Desenho & "'" Else TextoFiltro = "EE.ID_estoque = " & IDlista
CamposFiltro = "EE.ID, EE.Data as Dataemp, EE.Responsavel as Respemp, VC.CODIGO, EE.ID_estoque, Sum(EE.Qtde_empenhada) as qtdeliberar, Sum(EE.Qtde_saida) as qtdeliberada, VC.Unidade, VP.Ncotacao, VP.Revisao, OPCP.Requisicaotexto, VP.vend_int, VP.vend_ext, VC.Qtde_produzir, VC.qtdeexpedida, VC.Prazofinal, VC.Desenho, VC.Rev_codinterno, VC.N_Referencia, VC.descricao_tecnica, VP.Cliente"
CamposFiltroGrupo = "EE.ID, EE.Data, EE.Responsavel, VC.CODIGO, EE.ID_estoque, VC.Unidade, VP.Ncotacao, VP.Revisao, OPCP.Requisicaotexto, VP.vend_int, VP.vend_ext, VC.Qtde_produzir, VC.qtdeexpedida, VC.Prazofinal, VC.Desenho, VC.Rev_codinterno, VC.N_Referencia, VC.descricao_tecnica, VP.Cliente"

StrSql = "Select " & CamposFiltro & " from ((vendas_carteira VC INNER JOIN Estoque_Controle_Empenho_Vendas EE ON VC.Codigo = EE.ID_carteira " & IIf(Desenho <> "", "and VC.Desenho = '" & Desenho & "'", "") & ") LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OPCP ON OPCP.ID = VC.ID_Solicitacao group by " & CamposFiltroGrupo & " HAVING " & TextoFiltro & " and Sum(EE.Qtde_empenhada) - Sum(ISNULL(EE.Qtde_saida, 0)) > 0"
'Debug.print StrSql

StrSql = ""

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CamposFiltro & " from ((vendas_carteira VC INNER JOIN Estoque_Controle_Empenho_Vendas EE ON VC.Codigo = EE.ID_carteira " & IIf(Desenho <> "", "and VC.Desenho = '" & Desenho & "'", "") & ") LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OPCP ON OPCP.ID = VC.ID_Solicitacao group by " & CamposFiltroGrupo & " HAVING " & TextoFiltro & " and Sum(EE.Qtde_empenhada) - Sum(ISNULL(EE.Qtde_saida, 0)) > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaPedidos.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!CODIGO
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Dataemp), "", Format(TBLISTA!Dataemp, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Respemp), "", TBLISTA!Respemp)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!ID_estoque), "", TBLISTA!ID_estoque)
            If IsNull(TBLISTA!Ncotacao) = False Then
                .Item(.Count).SubItems(5) = TBLISTA!Ncotacao
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!vend_int), "", TBLISTA!vend_int)
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Vend_ext), "", TBLISTA!Vend_ext)
            Else
                .Item(.Count).SubItems(5) = TBLISTA!Requisicaotexto
            End If
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Qtde_produzir), "", Format(TBLISTA!Qtde_produzir, "###,##0.0000"))
            valor = IIf(IsNull(TBLISTA!qtdeliberar), 0, TBLISTA!qtdeliberar)
            .Item(.Count).SubItems(12) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!qtdeliberada), 0, TBLISTA!qtdeliberada)
            .Item(.Count).SubItems(13) = Format(Valor1, "###,##0.0000")
            .Item(.Count).SubItems(14) = Format(valor - Valor1, "###,##0.0000")
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
