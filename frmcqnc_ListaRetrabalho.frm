VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmcqnc_ListaRetrabalho 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NC - Lista retrabalho"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   4080
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7197
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   2020
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   2020
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2020
      EndProperty
   End
End
Attribute VB_Name = "frmcqnc_ListaRetrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select
            
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

With frmcqnc
    Lista.ListItems.Clear
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Ordem, IDProducao, Quantidade from Ordemservico where Ordem = " & .txtOF & " and fase = " & .txtFase & " and Retrabalho = 'True'", Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            With Lista.ListItems
                .Add , , TBLISTA!Ordem
                .Item(.Count).SubItems(1) = TBLISTA!IDProducao
                .Item(.Count).SubItems(2) = TBLISTA!quantidade
            End With
            TBLISTA.MoveNext
        Loop
    End If
    
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select P.Ordem, P.QUANT from CQ_NC_FABRICA C INNER JOIN Producao P ON C.OrdemRetrabalho = P.Ordem where C.codigo = " & .txtidos & " and C.OrdemRetrabalho IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            With Lista.ListItems
                .Add , , TBLISTA!Ordem
                .Item(.Count).SubItems(1) = ""
                .Item(.Count).SubItems(2) = TBLISTA!Quant
            End With
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
    
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Ordem, lote from CQ_NC_FABRICA where OrdemRetrabalho = " & .txtOF & " GROUP BY ordem, lote", Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            With Lista.ListItems
                .Add , , TBLISTA!Ordem
                .Item(.Count).SubItems(1) = ""
                .Item(.Count).SubItems(2) = TBLISTA!LOTE
            End With
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
