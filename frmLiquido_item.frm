VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmLiquido_item 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Ensaios - Líquido penetrante - Localizar produtos"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView ListaItem 
      Height          =   5535
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   9763
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
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Vlr. unitário"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Desc. (%)"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Vlr. desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Vlr. unit. c/ desc."
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   5550
      Width           =   10815
      _ExtentX        =   19076
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
Attribute VB_Name = "frmLiquido_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcCarregaLista()
On Error GoTo tratar_erro

ListaItem.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
If Liquido = True Then TBLISTA.Open "Select * from vendas_carteira where proposta = '" & frmLiquido.txtPedido_interno & "' and tipo = 'P' order by desenho", Conexao, adOpenKeyset, adLockOptimistic
If Ultrasom = True Then TBLISTA.Open "Select * from vendas_carteira where proposta = '" & frmUltraSom.txtPedido_interno & "' and tipo = 'P' order by desenho", Conexao, adOpenKeyset, adLockOptimistic
If Liquido = False And Ultrasom = False Then TBLISTA.Open "Select * from vendas_carteira where proposta = '" & frmCertificado_qualidade.txtPedido_interno & "' and tipo = 'P' order by desenho", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaItem.ListItems
            .Add , , TBLISTA!Desenho
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.0000"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!preco_unitario), "", Format(TBLISTA!preco_unitario, "###,##0.0000000000"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Desconto), "", TBLISTA!Desconto)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!ValorDesconto), "", Format(TBLISTA!ValorDesconto, "###,##0.0000000000"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!preco_unitario_desconto), "", Format(TBLISTA!preco_unitario_desconto, "###,##0.0000000000"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!preco_lote), "", Format(TBLISTA!preco_lote, "###,##0.0000000000"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Liberacao), "", TBLISTA!Liberacao)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: ListaItem_DblClick
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Ultrasom = True Then frmLiquido_item.Caption = "Controle da qualidade - Ensaio por ultra-som - Localizar produtos"
If Liquido = True Then frmLiquido_item.Caption = "Controle da qualidade - Ensaio por líquido penetrante - Localizar produtos"
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ListaItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaItem, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaItem_DblClick()
On Error GoTo tratar_erro

If ListaItem.ListItems.Count = 0 Then Exit Sub
If Liquido = True Then
    With frmLiquido
        .txtDesenho = ListaItem.SelectedItem
        .txtDescricao = ListaItem.SelectedItem.SubItems(1)
    End With
End If
If Ultrasom = True Then
    With frmUltraSom
        .txtDesenho = ListaItem.SelectedItem
        .txtDescricao = ListaItem.SelectedItem.SubItems(1)
    End With
End If
If Ultrasom = False And Liquido = False Then
    With frmCertificado_qualidade
        .txtDesenho = ListaItem.SelectedItem
        .txtDescricao = ListaItem.SelectedItem.SubItems(1)
        .ProcCarregalista_ultra
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
