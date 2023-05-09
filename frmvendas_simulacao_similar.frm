VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmvendas_simulacao_similar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Vendas - Simulação - Produtos similiares"
   ClientHeight    =   6360
   ClientLeft      =   1680
   ClientTop       =   1365
   ClientWidth     =   10200
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista 
      Height          =   6000
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   10090
      _ExtentX        =   17806
      _ExtentY        =   10583
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Qtde. est."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. custo"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   6060
      Width           =   10090
      _ExtentX        =   17806
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   1
      Top             =   1320
      Width           =   45
   End
End
Attribute VB_Name = "frmvendas_simulacao_similar"
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
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select P.Codproduto, P.Desenho, P.Descricao, P.Unidade, P.PCusto, IA.n_referencia from Projproduto P LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto where P.ID_similar = " & IDlista & " and P.Codproduto <> " & frmvendas_simulacao.IDProduto, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
       With Lista.ListItems
            .Add , , TBLISTA!Codproduto
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!N_referencia), "", TBLISTA!N_referencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(5) = Format(Qtd_Prog, "###,##0.0000")
            .Item(.Count).SubItems(6) = Format(FunVerificaQtdeEstoque(TBLISTA!Desenho, 0, ""), "###,##0.0000")
            
            If frmvendas_simulacao.Chk_valor.Value = 1 Then
                'Verifica custo médido do estoque sem material consignado
                Call FunVerificaQtdeEstoque(TBLISTA!Desenho, 0, "and Consignacao = 'False'")
                If CTMedioEst <> 0 Then
                    valor = Format(CTMedioEst * Qtd_Prog, "###,##0.00000000")
                Else
                    'Verifica valor unitário da última compra
                    valor = FunVerificaVlrUltCompra(TBLISTA!Desenho)
                    If valor <> 0 Then
                        valor = Format(valor * Qtd_Prog, "###,##0.00000000")
                    Else
                        valor = Format(IIf(IsNull(TBLISTA!PCusto), 0, TBLISTA!PCusto * Qtd_Prog), "###,##0.00000000")
                    End If
                End If
            Else
                valor = 0
            End If
            .Item(.Count).SubItems(7) = Format(valor, "###,##0.00000000")
            
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
