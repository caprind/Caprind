VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Orcamento_Terceiros 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Vendas | Orçamento | Terceiros"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   6945
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   873
      DibPicture      =   "frm_Orcamento_Terceiros.frx":0000
      Caption         =   "Vendas | Orçamento | Itens"
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
      Icon            =   "frm_Orcamento_Terceiros.frx":1CAD
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Valor total"
      Top             =   600
      Width           =   8115
      Begin VB.ComboBox cmbcodref 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frm_Orcamento_Terceiros.frx":1FC7
         Left            =   1470
         List            =   "frm_Orcamento_Terceiros.frx":1FC9
         Sorted          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Código de referência."
         Top             =   435
         Width           =   1560
      End
      Begin VB.TextBox cmbfamilia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Família."
         Top             =   1590
         Width           =   6675
      End
      Begin VB.TextBox txtvlrTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtvlrUnitario 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Valor unitario do item"
         Top             =   450
         Width           =   1035
      End
      Begin VB.TextBox txtcodigoproduto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtunidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   450
         Width           =   555
      End
      Begin VB.TextBox txtdescricao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Width           =   6690
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4650
         TabIndex        =   3
         ToolTipText     =   "Quantidade"
         Top             =   450
         Width           =   885
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Código"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   4
         Left            =   3195
         TabIndex        =   19
         Top             =   240
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   344
         Caption         =   "Un"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Un"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   5
         Left            =   3113
         TabIndex        =   20
         Top             =   810
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   "Descrição"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Descrição"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   0
         Left            =   3645
         TabIndex        =   21
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   344
         Caption         =   "Valor Unitário"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Valor Unitário"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   1
         Left            =   5790
         TabIndex        =   22
         Top             =   240
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   344
         Caption         =   "valor Total"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "valor Total"
      End
      Begin DrawSuite2022.USButton btnNovo 
         Height          =   555
         Left            =   6990
         TabIndex        =   11
         Top             =   180
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         DibPicture      =   "frm_Orcamento_Terceiros.frx":1FCB
         Caption         =   "Novo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   8
         PicSize         =   1
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton btnGravar 
         Height          =   555
         Left            =   6990
         TabIndex        =   12
         Top             =   750
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         DibPicture      =   "frm_Orcamento_Terceiros.frx":82AF
         Caption         =   "Gravar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   8
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton btnExcluir 
         Height          =   585
         Left            =   6990
         TabIndex        =   13
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1032
         DibPicture      =   "frm_Orcamento_Terceiros.frx":10CB4
         Caption         =   "Excluir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicAlign        =   8
         ShowFocusRect   =   0   'False
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Família"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3277
         TabIndex        =   15
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4065
      Left            =   150
      TabIndex        =   10
      Top             =   2730
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7170
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "T"
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Codigo"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Un"
         Object.Width           =   705
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5847
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Quant"
         Object.Width           =   1146
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Valor"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frm_Orcamento_Terceiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ID_Conjunto As Integer

Private Sub btnExcluir_Click()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
    If USMsgBox("Deseja realmente excluir o item " & Lista.SelectedItem.ListSubItems.Item(1).Text & "?", vbYesNo, "CAPRIND  v5.0") = vbYes Then
        Conexao.Execute ("Delete from Vendas_Orcamento_Terceiros where id = '" & Lista.SelectedItem & "'")
        USMsgBox "Item excluido com sucesso!", vbInformation, "CAPRIND v5.0"
        ProcCarregaLista_Terceiros
    End If
End If

ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnGravar_Click()
On Error GoTo tratar_erro

ProcgravarItem

Set TBFases = CreateObject("adodb.recordset")
StrSql = "Select SUM(vlrtotal) as Total from Vendas_Orcamento_Terceiros where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBFases.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

    If TBFases.EOF = False Then
        frm_orcamento.txtv3.Text = Format(TBFases!Total, "###,##0.00")
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcgravarItem()
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
StrSql = "Select * from Vendas_Orcamento_Terceiros where ID_Orcamento = '" & frm_orcamento.txtId.Text & "' and ID = '" & ID_Terceiro & "'"
TBItem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = True Then
TBItem.AddNew
End If
TBItem!ID_orcamento = frm_orcamento.txtId
TBItem!Codproduto = Cod_produto
TBItem!CODIGO = txtcodigoproduto.Text
TBItem!Unidade = txtunidade.Text
TBItem!Descricao = txtdescricao.Text
TBItem!quantidade = txtLote.Text
TBItem!vlrUnitario = Format((txtvlrTotal.Text / txtLote), "###,##0.0000")
TBItem!vlrTotal = Format(txtvlrTotal.Text, "###,##0.0000")
TBItem.Update
TBItem.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND  V5.0"

Conexao.Execute (StrSql)

ProcCarregaLista_Terceiros

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_Terceiros()
On Error GoTo tratar_erro

valor = 0
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select * from Vendas_Orcamento_Terceiros where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!quantidade), "", TBLISTA!quantidade)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!vlrUnitario), "", "R$ " & Format(TBLISTA!vlrUnitario, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!vlrTotal), "", "R$ " & Format(TBLISTA!vlrTotal, "###,##0.00"))
            valor = valor + IIf(IsNull(TBLISTA!vlrTotal), 0, TBLISTA!vlrTotal)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
    Lista.ListItems.Add , , 1 'TBLISTA!ID_Conjunto
    Lista.ListItems.Item(Contador + 1).SubItems(5) = "TOTAL :"
    Lista.ListItems.Item(Contador + 1).ListSubItems.Item(4).ForeColor = vbRed
    Lista.ListItems.Item(Contador + 1).SubItems(6) = "R$ " & Format(valor, "###,##0.00")
    Lista.ListItems.Item(Contador + 1).ListSubItems.Item(6).ForeColor = vbRed
    

frm_orcamento.txtv3.Text = Format(valor, "###,##0.00")

End If
TBLISTA.Close

        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnNovo_Click()
On Error GoTo tratar_erro

ProcLimparCampos
frm_Orcamento_Terceiros_item_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcLimparCampos

If frm_orcamento.txtcodproduto <> "" Then
    ProcCarregaLista_Terceiros
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaValor()
On Error GoTo tratar_erro


If txtvlrUnitario.Text <> "" And txtLote.Text <> "" Then
 txtvlrTotal = Format(txtvlrUnitario * txtLote, "###,##0.00")
End If
    

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro
      
      ID_Conjunto = 0
      Cod_produto = 0
      txtcodigoproduto = ""
      txtdescricao.Text = ""
      txtLote = ""
      txtunidade = ""
      txtvlrUnitario.Text = ""
      txtvlrTotal.Text = ""
      cmbfamilia.Text = ""
      txtvlrTotal = Format(0, "###,##0.00")
      cmbfamilia = ""


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
  If Lista.SelectedItem <> "" Then
  
  Set TBLISTA = CreateObject("adodb.recordset")
  
  StrSql = "select VOC.*, PP.Classe,PP.codproduto,PP.Un_Kg from Vendas_Orcamento_Terceiros VOC inner join projproduto PP on PP.codproduto = VOC.codProduto where VOC.ID = '" & Lista.SelectedItem & "'"
  'Debug.print StrSql
  
  TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBLISTA.EOF = False Then
    ID_Conjunto = Lista.SelectedItem
      Cod_produto = TBLISTA!Codproduto
      txtcodigoproduto.Text = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
      txtdescricao.Text = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
      txtvlrUnitario = IIf(IsNull(TBLISTA!vlrUnitario), "", Format(TBLISTA!vlrUnitario, "###,##0.00"))
      txtLote.Text = IIf(IsNull(TBLISTA!quantidade), "", Format(TBLISTA!quantidade, "###,##0.00"))
      txtunidade.Text = IIf(IsNull(TBLISTA!Unidade), "", TBLISTA!Unidade)
      cmbfamilia.Text = IIf(IsNull(TBLISTA!Classe), "", TBLISTA!Classe)
      txtvlrTotal = IIf(IsNull(TBLISTA!vlrTotal), "", Format(TBLISTA!vlrTotal, "###,##0.00"))
  End If
  TBLISTA.Close
  
  End If
End If

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from item_aplicacoes where codproduto = " & Cod_produto & "", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
 cmbcodref.Text = TBItem!N_referencia
End If
TBItem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLote_Change()
On Error GoTo tratar_erro

If txtLote.Text <> "" Then
    VerifNumero = txtLote.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLote.Text = ""
        txtLote.SetFocus
        Exit Sub
    End If
End If

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
