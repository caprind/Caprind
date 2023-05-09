VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_propostaII_ListaAnalise 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Vendas - Proposta comercial  - Lista de análise crítica"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendas_propostaII_ListaAnalise.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   4365
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7699
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
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
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Análise"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Descição"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   4380
      Width           =   8685
      _ExtentX        =   15319
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
Attribute VB_Name = "frmVendas_propostaII_ListaAnalise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
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

If Vendas_PI = True Then Caption = "Administrativo - Vendas - Pedido interno  - Lista de análise crítica"
With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    If Vendas_Produtos = True Then
        If .txtNomenclatura = "" And .txtRev_cod = "" Then
            TextoFiltro = "IDCliente = " & .txtIDCliente
        ElseIf .txtNomenclatura <> "" And .txtRev_cod = "" Then
                TextoFiltro = "Codinterno = '" & .txtNomenclatura & "' and IDCliente = " & .txtIDCliente
            Else
                TextoFiltro = "Codinterno = '" & .txtNomenclatura & "' and Revdesenho = '" & .txtRev_cod & "' and IDCliente = " & .txtIDCliente
        End If
    Else
        If .txtcodservico = "" And .txtRev_serv = "" Then
            TextoFiltro = "IDCliente = " & .txtIDCliente
        ElseIf .txtcodservico <> "" And .txtRev_serv = "" Then
                TextoFiltro = "Codinterno = '" & .txtcodservico & "' and IDCliente = " & .txtIDCliente
            Else
                TextoFiltro = "Codinterno = '" & .txtcodservico & "' and Revdesenho = '" & .txtRev_serv & "' and IDCliente = " & .txtIDCliente
        End If
    End If
End With

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Vendas_analise where " & TextoFiltro & " and Status = 'APROVADA' order by Ordenaranalise desc, ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Nanalise), "", TBLISTA!Nanalise)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!RevDesenho), "", TBLISTA!RevDesenho)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Valor_total), "", Format(TBLISTA!Valor_total, "###,##0.0000000000"))
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
Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select ID, Codinterno, RevDesenho, N_Referencia, Descricao, Unidade, Unidade_com, Familia, Qtde from Vendas_analise where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        If Vendas_Produtos = True Then
            .txtNomenclatura = IIf(IsNull(TBTempo!Codinterno), "", TBTempo!Codinterno)
            .txtRev_cod = IIf(IsNull(TBTempo!RevDesenho), "", TBTempo!RevDesenho)
            .cmbreferencia.Clear
            If IsNull(TBTempo!N_referencia) = False And TBTempo!N_referencia <> "" Then
                .cmbreferencia.AddItem TBTempo!N_referencia
                .cmbreferencia = TBTempo!N_referencia
            End If
            .txtdesctecnica = IIf(IsNull(TBTempo!Descricao), "", TBTempo!Descricao)
            .IDAnalise = TBTempo!ID
            .Txt_analise = Lista.SelectedItem.ListSubItems(1)
            .txtEspecificacoes = IIf(IsNull(TBTempo!Descricao), "", TBTempo!Descricao)
            .cmbun = IIf(IsNull(TBTempo!Unidade), "", TBTempo!Unidade)
            .Cmb_un_com = IIf(IsNull(TBTempo!Unidade_com), "", TBTempo!Unidade_com)
            .cmbfamilia = IIf(IsNull(TBTempo!Familia), "", TBTempo!Familia)
            .txtQuantidade = IIf(IsNull(TBTempo!Qtde), "", Format(TBTempo!Qtde, "###,##0.0000"))
            .txtvalorunitario = Lista.SelectedItem.ListSubItems(6)
        Else
            .txtcodservico = IIf(IsNull(TBTempo!Codinterno), "", TBTempo!Codinterno)
            .txtRev_serv = IIf(IsNull(TBTempo!RevDesenho), "", TBTempo!RevDesenho)
            .cmbreferencia_serv.Clear
            If IsNull(TBTempo!N_referencia) = False And TBTempo!N_referencia <> "" Then
                .cmbreferencia_serv.AddItem TBTempo!N_referencia
                .cmbreferencia_serv = TBTempo!N_referencia
            End If
            .txtdescservico = IIf(IsNull(TBTempo!Descricao), "", TBTempo!Descricao)
            .IDAnalise_servico = TBTempo!ID
            .Txt_analise1 = Lista.SelectedItem.ListSubItems(1)
            .txtdesccomservico = IIf(IsNull(TBTempo!Descricao), "", TBTempo!Descricao)
            .cmbfamiliaservico = IIf(IsNull(TBTempo!Familia), "", TBTempo!Familia)
            .txtunservico = IIf(IsNull(TBTempo!Unidade), "", TBTempo!Unidade)
            .Cmb_un_com_serv = IIf(IsNull(TBTempo!Unidade_com), "", TBTempo!Unidade_com)
            .txtqtservico = IIf(IsNull(TBTempo!Qtde), "", Format(TBTempo!Qtde, "###,##0.0000"))
            .txtvlrunitservico = Lista.SelectedItem.ListSubItems(6)
        End If
    End With
End If
TBTempo.Close
Unload Me
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
