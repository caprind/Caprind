VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmVendas_analise_CopiaEngenharia 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Análise crítica - Copiar dados da engenharia"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
         Text            =   "Análise"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Rev."
         Object.Width           =   970
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
Attribute VB_Name = "frmVendas_analise_CopiaEngenharia"
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
    TBLISTA.Open "Select * from Vendas_analise where ID <> " & .TxtID & " order by Ordenaranalise desc, ID desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        TBLISTA.MoveLast
        PBLista.Min = 0
        PBLista.Max = TBLISTA.RecordCount
        PBLista.Value = 1
        Contador = 0
        TBLISTA.MoveFirst
        Do While TBLISTA.EOF = False
            With Lista.ListItems
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "select * from vendas_analise_setores where IDanalise = " & TBLISTA!ID & " and Setor = 'ENGENHARIA'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Add , , TBLISTA!ID
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Nanalise), "", TBLISTA!Nanalise)
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Codinterno), "", TBLISTA!Codinterno)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!RevDesenho), "", TBLISTA!RevDesenho)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
                End If
                TBAbrir.Close
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

Private Sub imgSair_Click()
On Error GoTo tratar_erro
  
Unload Me

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
    TBAbrir.Open "Select * from vendas_analise_setores where IDanalise = " & Lista.SelectedItem & " and Setor = 'ENGENHARIA'", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBAbrir.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from vendas_analise_setores", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!IDAnalise = .TxtID
        TBGravar!Responsavel = pubUsuario
        TBGravar!data = Date
        TBGravar!IDProduto = IIf(IsNull(TBAbrir!IDProduto), "", TBAbrir!IDProduto)
        TBGravar!Texto = IIf(IsNull(TBAbrir!Texto), "", TBAbrir!Texto)
        TBGravar!Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
        TBGravar!Analise = IIf(IsNull(TBAbrir!Analise), "", TBAbrir!Analise)
        TBGravar!Familia = IIf(IsNull(TBAbrir!Familia), "", TBAbrir!Familia)
        TBGravar!Un = IIf(IsNull(TBAbrir!Un), "", TBAbrir!Un)
        TBGravar!Unidade_com = IIf(IsNull(TBAbrir!Unidade_com), "", TBAbrir!Unidade_com)
        TBGravar!N_referencia = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
        TBGravar!Codinterno = IIf(IsNull(TBAbrir!Codinterno), "", TBAbrir!Codinterno)
        TBGravar!Tipo = IIf(IsNull(TBAbrir!Tipo), "", TBAbrir!Tipo)
        TBGravar!Setor = "ENGENHARIA"
        TBGravar.Update
        TBAbrir.MoveNext
    Loop
    TBAbrir.Close
    
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Copiar engenharia"
    ID_documento = .TxtID.Text
    Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
    USMsgBox ("Material/terceiro(s) da engenharia copiada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
End With
Unload Me
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
