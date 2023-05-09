VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProd_programacao_OS 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Programação da produção - Localizar OS"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12135
   ClipControls    =   0   'False
   Icon            =   "frmProd_programacao_OS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmProd_programacao_OS.frx":1042
   MousePointer    =   99  'Custom
   ScaleHeight     =   6990
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lista 
      Height          =   6945
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   12250
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmProd_programacao_OS.frx":134C
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   2593
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Grupo/op."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Codigo interno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmProd_programacao_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Ordem_programacao_LocalizarOS As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaLista

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
With frmProd_programacao
    .txtOS = Lista.SelectedItem
    .ProcCarregaOS_Fase
End With
Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from ordemservico where pronto = 'NÃO' and maquina = '" & frmProd_programacao.cmbMaquina & "' order by idproducao desc", Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from producao where Ordem = " & TBLISTA!Ordem & "", Conexao, adOpenKeyset, adLockOptimistic
        
        With Lista.ListItems
            .Add , , TBLISTA!IDProducao
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Grupo_op), "", TBLISTA!Grupo_op)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBOrdem!Desenho), "", TBOrdem!Desenho)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBOrdem!Produto), "", TBOrdem!Produto)
            
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
