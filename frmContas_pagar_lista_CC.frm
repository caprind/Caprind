VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmContas_pagar_lista_CC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Contas a pagar - Centro de custo"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView Lista 
      Height          =   4110
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   7250
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9657
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   4155
      Width           =   8565
      _ExtentX        =   15108
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
Attribute VB_Name = "frmContas_pagar_lista_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

If Financeiro_Contas_Pagas = True Then Caption = "Administrativo - Financeiro - Contas pagas - Centro de custo"
ProcLimpaVariaveisPrincipais
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBAbrir = CreateObject("adodb.recordset")
Set TBLISTA = CreateObject("adodb.recordset")
If IsNull(TBContas!ID_nota) = False And TBContas!ID_nota <> "" And TBContas!ID_nota <> "0" Or IsNull(TBContas!txt_ndocumento) = False And TBContas!txt_ndocumento <> "" Then
    INNERJOINTEXTO = "(tbl_Detalhes_Nota NFP INNER JOIN Compras_pedido_lista CPL ON CPL.IDlista = NFP.Codigo and CPL.Desenho = NFP.int_Cod_Produto) INNER JOIN Compras_pedido_lista_custo CPLC ON CPLC.IDlista = CPL.IDlista"
    TextoFiltro = "NFP.ID_nota = " & TBContas!ID_nota
    TBLISTA.Open "Select CPLC.ID, CPLC.Valor, US.Codigo, US.Setor from " & INNERJOINTEXTO & " INNER JOIN Usuarios_Setor US ON US.ID = CPLC.ID_CC where " & TextoFiltro & " Order by US.Codigo", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBLISTA.Open "Select CPLC.ID, CPLC.Valor, US.Codigo, US.Setor from ((Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDpedido = CPL.IDPedido) INNER JOIN Compras_pedido_lista_custo ON CPLC CPL.IDlista = CPLC.IDlista) INNER JOIN Usuarios_Setor US ON US.ID = CPLC.ID_CC where CP.Pedido = '" & TBContas!Txt_pedido & "' Order by US.Codigo", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "0,00", Format(TBLISTA!valor, "###,##0.00"))
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
