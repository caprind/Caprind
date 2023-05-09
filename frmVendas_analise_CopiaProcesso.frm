VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_analise_CopiaProcesso 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Análise crítica - Copiar fase(s) do processo"
   ClientHeight    =   4725
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
   Icon            =   "frmVendas_analise_CopiaProcesso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4725
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   4380
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "IDproc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Análise"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descição"
         Object.Width           =   8291
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   4440
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmVendas_analise_CopiaProcesso"
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

With frmVendas_analise
    Lista.ListItems.Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select VA.ID, VA.Nanalise, VA.Revisao, VAP.ID AS IDprodproc, VAP.Codinterno, VAP.Descricao from (Vendas_analise VA INNER JOIN Vendas_analise_ProdutosProcessos VAP ON VA.ID = VAP.ID_analise) INNER JOIN vendas_analise_setores VAS ON VA.ID = VAS.IDanalise where VA.ID <> " & .TxtID & " and VAS.Setor = 'PROCESSOS' and VAP.Codinterno = '" & .txtCodInterno_processos_item & "' group by VA.Ordenaranalise, VA.ID, VA.Nanalise, VA.Revisao, VAP.ID, VAP.Codinterno, VAP.Descricao order by VA.Ordenaranalise desc, VA.ID desc", Conexao, adOpenKeyset, adLockReadOnly
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
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDprodproc), "", TBLISTA!IDprodproc)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Nanalise), "", TBLISTA!Nanalise)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Revisao), 0, TBLISTA!Revisao)
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            End With
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End If
    TBLISTA.Close
End With
    
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
With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise_setores where ID_processo_item = " & Lista.SelectedItem.ListSubItems(1) & " and Setor = 'PROCESSOS' order by Fase", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBAbrir.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from vendas_analise_setores", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        
        'Salva e verifica se repete a fase na mesma analize
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select * from vendas_analise_setores where ID_processo_item = " & .txtID_processos_item & " and Setor = 'PROCESSOS' and Fase = " & TBAbrir!Fase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            TBGravar!Fase = TBAbrir!Fase + 1
1:
            Set TBProcessos = CreateObject("adodb.recordset")
            TBProcessos.Open "Select * from vendas_analise_setores where ID_processo_item = " & .txtID_processos_item & " and Setor = 'PROCESSOS' and Fase = " & TBGravar!Fase, Conexao, adOpenKeyset, adLockOptimistic
            If TBProcessos.EOF = False Then
                TBGravar!Fase = TBGravar!Fase + 1
                GoTo 1
            End If
            TBProcessos.Close
        Else
            TBGravar!Fase = TBAbrir!Fase
        End If
        TBFases.Close
        
        TBGravar!IDAnalise = .TxtID
        TBGravar!Responsavel = pubUsuario
        TBGravar!data = Date
        TBGravar!Texto = IIf(IsNull(TBAbrir!Texto), "", TBAbrir!Texto)
        TBGravar!Descricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
        TBGravar!Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
        TBGravar!Peca = IIf(IsNull(TBAbrir!Peca), 0, TBAbrir!Peca)
        TBGravar!Execucao = IIf(IsNull(TBAbrir!Execucao), Null, TBAbrir!Execucao)
        TBGravar!Preparacao = IIf(IsNull(TBAbrir!Preparacao), Null, TBAbrir!Preparacao)
        TBGravar!VlrUnit = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
        TBGravar!PrecoHora_Setup = IIf(IsNull(TBAbrir!PrecoHora_Setup), 0, TBAbrir!PrecoHora_Setup)
        TBGravar!Leadtime = TBAbrir!Leadtime
        TBGravar!Setor = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
        TBGravar!ID_processo_item = .txtID_processos_item
        If TBAbrir!pecahora = True Then
            TBGravar!pecahora = True
        Else
            TBGravar!pecahora = False
        End If
        TBGravar!TotalHora = IIf(IsNull(TBAbrir!TotalHora), Null, TBAbrir!TotalHora)
        TBGravar!Trabalho = IIf(IsNull(TBAbrir!Trabalho), "", TBAbrir!Trabalho)
        TBGravar!Grupo_op = IIf(IsNull(TBAbrir!Grupo_op), "", TBAbrir!Grupo_op)
        TBGravar!Erro_processos = IIf(IsNull(TBAbrir!Erro_processos), 0, TBAbrir!Erro_processos)
        
        'Calcula custo de preparação diluido na quantidade
        Qtde = 0
        valor = 0
        qt = 0
        Qtd = 0
        ValorTotal = 0
        quantidade = 0
        Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
        valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
            
        dataCalculo = IIf(IsNull(TBAbrir!Preparacao), 0, TBAbrir!Preparacao)
        ProcFormataHora (dataCalculo)
        qt = s / 3600
        Qtd = IIf(.txtQtde_processos_item = "", 0, .txtQtde_processos_item)
        ValorTotal = Qtde * valor
        If qt > 0 Then
            quantidade = (qt / Qtd) * valor
            ValorTotal = ValorTotal + quantidade
        End If
        TBGravar!VlrTotal = ValorTotal
        
        TBGravar.Update
        TBAbrir.MoveNext
    Loop
    TBAbrir.Close
    
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Copiar processos"
    ID_documento = .TxtID.Text
    Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
    USMsgBox ("Fase(s) do processo copiada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
End With
Unload Me
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
