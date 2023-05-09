VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargaMaq_ListaOrdem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Carga de posto de trabalho - Lista de ordem"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCargaMaq_ListaOrdem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   8190
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCargaMaq_ListaOrdem.frx":1042
      Count           =   1
   End
   Begin MSComctlLib.ListView Listaordem 
      Height          =   6345
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   11192
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Pr. final"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4119
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "D"
         Text            =   "Tempo total prev."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Tempo total prod."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "D"
         Text            =   "Tempo total rest."
         Object.Width           =   2293
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   12825
      _ExtentX        =   22622
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
      ButtonCaption2  =   "Concluir OS"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Marcar OS como concluída (F7)"
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
      ButtonLeft2     =   55
      ButtonTop2      =   2
      ButtonWidth2    =   63
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
      ButtonLeft3     =   120
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
      ButtonLeft4     =   124
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
      ButtonLeft5     =   162
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
      ButtonLeft6     =   190
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   7350
      Width           =   12825
      _ExtentX        =   22622
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
End
Attribute VB_Name = "frmCargaMaq_ListaOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcConcluirOS()
On Error GoTo tratar_erro

Permitido = False
With Listaordem
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente marcar esta(s) OS('s) como concluída?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBOrdemServico = CreateObject("adodb.recordset")
            TBOrdemServico.Open "Select P.Ap_backup, OS.* from producao P INNER JOIN ordemservico OS ON P.ordem = OS.Ordem where OS.IDPRODUCAO = " & .ListItems.Item(InitFor).ListSubItems(3).Text, Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdemServico.EOF = False Then
                If TBOrdemServico!AP_backup = True Then
                    NomeTabelaAp = "ProducaoFases_Backup"
                    NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao_Backup"
                Else
                    NomeTabelaAp = "ProducaoFases"
                    NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao"
                End If
                
                Conexao.Execute "Update " & NomeTabelaAp & " Set pronto = 'SIM' where idfase = " & TBOrdemServico!IDProducao
                
                TBOrdemServico!Pronto = "SIM"
                TBOrdemServico!DataConclusao = Date
                TBOrdemServico!status = "Concluída"
                TBOrdemServico.Update
                
                Conexao.Execute "Update Ordemservico_maq_utilizadas Set Pronto = 'SIM' where OS = " & TBOrdemServico!IDProducao
                Conexao.Execute "Update CM Set CM.Liberada = 'True' from cadmaquinas CM INNER JOIN cadmaquinas_Monitor CMM ON CM.Maquina = CMM.Maquina where CMM.maquina = '" & TBOrdemServico!maquina & "' and CMM.OS = " & TBOrdemServico!IDProducao
                
                ProcConcluirOrdem TBOrdemServico!Ordem
                '==================================
                Modulo = "PCP/Carga de posto de trabalho"
                Evento = "Alterar OS p/ concluída"
                ID_documento = .ListItems.Item(InitFor).ListSubItems(3).Text
                Documento = "Ordem: " & OF & " - Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(4).Text
                Documento1 = "OS: " & .ListItems.Item(InitFor).ListSubItems(3).Text
                ProcGravaEvento
                '==================================
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from ordemservico where Ordem = " & .ListItems(InitFor).ListSubItems(2) & " and ID_apontamento is not null order by Idproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from ordemservico where IDproducao <> " & TBAbrir!IDProducao & " and ID_apontamento = " & TBAbrir!ID_apontamento, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFI.EOF = False Then
                            If TBFI.RecordCount <= 1 Then Conexao.Execute "DELETE from ProducaoFases_OS WHERE ID = " & TBAbrir!ID_apontamento
                        Else
                            Conexao.Execute "DELETE from ProducaoFases_OS WHERE ID = " & TBAbrir!ID_apontamento
                        End If
                        TBFI.Close
                        TBAbrir.MoveNext
                    Loop
                End If
                Conexao.Execute "DELETE from ordemservico where Ordem = " & .ListItems(InitFor).ListSubItems(2)
                Conexao.Execute "DELETE from Ordemservico_maq_utilizadas where Ordem = " & .ListItems(InitFor).ListSubItems(2)
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) OS('s) na lista antes de marcar como concluída."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcConcluirOrdem(OF As Long)
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    'Checa se todas as ordens de servicos com processo da ordem foram executadas e da baixa na ordem
    Set TBOrdemServico = CreateObject("adodb.recordset")
    TBOrdemServico.Open "Select * from ordemservico where Ordem = " & OF & " and pronto = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdemServico.EOF = True Then
        TBAbrir!pronta = "SIM"
        TBAbrir!Concluida = True
        If TBAbrir!status <> "Entregue" Then TBAbrir!status = "Concluída"
        TBAbrir!DataEntrega = Date
    Else
        TBAbrir!pronta = "NÃO"
        TBAbrir!Concluida = False
        If TBAbrir!status <> "Cancelada" And TBAbrir!status <> "Aguardando" And TBAbrir!status <> "Entregue" And TBAbrir!status <> "Sem material" Then
            Set TBProducaoFases = CreateObject("adodb.recordset")
            TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
            If TBProducaoFases.EOF = False Then
                TBAbrir!status = "Produzindo"
            Else
                TBAbrir!status = "Aberta"
            End If
            TBProducaoFases.Close
        End If
        TBAbrir!DataEntrega = Null
    End If
    TBOrdemServico.Close
    TBAbrir.Update

    'Verifica se todas as ordems de fabricação do produto já foram concluidas
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from producao_pedidos where Ordem = " & OF & " order by IDCarteira", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select * from Vendas_carteira where Codigo = " & TBFIltro!IDcarteira, Conexao, adOpenKeyset, adLockOptimistic
            If TBVendas.EOF = False Then
                Set TBProcessos = CreateObject("adodb.recordset")
                TBProcessos.Open "Select producao.* FROM Producao INNER JOIN Producao_pedidos ON Producao.Ordem = Producao_pedidos.Ordem where Producao_pedidos.IDCarteira = " & TBFIltro!IDcarteira & " and Producao.pronta = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessos.EOF = True Then
                    TBVendas!saida_estoque = True
                    TBVendas!dataprodsaida = Date
                Else
                    TBVendas!saida_estoque = False
                    TBVendas!dataprodsaida = Null
                End If
                TBProcessos.Close
                TBVendas.Update
            End If
            TBVendas.Close
            TBFIltro.MoveNext
        Loop
    End If
    TBFIltro.Close
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcConcluirOS
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
        
ProcCarregaToolBar1 Me, 12825, 6, True
ProcCarregaLista
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro
        
Listaordem.ListItems.Clear
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select OS.IDProducao, OS.Prazofinal, OS.Ordem, OS.IDProducao, P.Desenho, OS.Quantidade, OS.Tempototallote, OS.TETTUTILSEG, OS.TTLPREVS, P.N_Referencia, P.Produto from ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where P.Status <> 'Cancelada' and OS.maquina = '" & frmCargaMaq.Lista.SelectedItem & "' AND OS.Prazofinal <= '" & Format(frmCargaMaq.msk_fltFim.Value, "Short Date") & "' and OS.pronto = 'Não' order by OS.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    TBproducao.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBproducao.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBproducao.MoveFirst
    Do While TBproducao.EOF = False
        With Listaordem.ListItems
            .Add , , TBproducao!IDProducao
            .Item(.Count).SubItems(1) = Format(TBproducao!PrazoFinal, "dd/mm/yy")
            .Item(.Count).SubItems(2) = TBproducao!Ordem
            .Item(.Count).SubItems(3) = TBproducao!IDProducao
            .Item(.Count).SubItems(4) = IIf(IsNull(TBproducao!Desenho), "", TBproducao!Desenho)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBproducao!N_referencia), "", TBproducao!N_referencia)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBproducao!Produto), "", TBproducao!Produto)
            .Item(.Count).SubItems(7) = Format(TBproducao!quantidade, "###,##0.00")
            .Item(.Count).SubItems(8) = IIf(IsNull(TBproducao!TempoTotalLote), "00:00:00", TBproducao!TempoTotalLote)
            .Item(.Count).SubItems(9) = FormataTempo(IIf(IsNull(TBproducao!TETTUTILSEG), 0, TBproducao!TETTUTILSEG))
            .Item(.Count).SubItems(10) = FormataTempo(IIf(IIf(IsNull(TBproducao!TTLPREVS), 0, TBproducao!TTLPREVS) - IIf(IsNull(TBproducao!TETTUTILSEG), 0, TBproducao!TETTUTILSEG) < 0, 0, IIf(IsNull(TBproducao!TTLPREVS), 0, TBproducao!TTLPREVS) - IIf(IsNull(TBproducao!TETTUTILSEG), 0, TBproducao!TETTUTILSEG)))
        End With
        TBproducao.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBproducao.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

NomeRel = "Pcp_cargapostodetrabalho_listaOrdem.rpt"
TBproducao.Open "Select OS.IDProducao, OS.Prazofinal, OS.Ordem, OS.IDProducao, P.Desenho, OS.Quantidade, OS.Tempototallote, OS.TETTUTILSEG, OS.TTLPREVS, P.N_Referencia, P.Produto from ordemservico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where P.Status <> 'Cancelada' and OS.maquina = '" & frmCargaMaq.Lista.SelectedItem & "' AND OS.Prazofinal <= '" & Format(frmCargaMaq.msk_fltFim.Value, "Short Date") & "' and OS.pronto = 'Não' order by OS.Prazofinal", Conexao, adOpenKeyset, adLockOptimistic
ProcImprimirRel "{Producao.Status} <> 'Cancelada' and {ordemservico.maquina} = '" & frmCargaMaq.Lista.SelectedItem & "' AND {ordemservico.Prazofinal} <= Date(" & Year(frmCargaMaq.msk_fltFim.Value) & "," & Month(frmCargaMaq.msk_fltFim.Value) & "," & Day(frmCargaMaq.msk_fltFim.Value) & ") and {ordemservico.pronto} = 'NÃO'", ""
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaordem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Listaordem
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Listaordem, ColumnHeader
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
    Case 2: ProcConcluirOS
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
