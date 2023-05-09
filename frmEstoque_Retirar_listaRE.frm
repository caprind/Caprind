VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmEstoque_Retirar_listaRE 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Estoque | Retirar - Lista de RE´s"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4260
   StartUpPosition =   3  'Padrão Windows
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   2730
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      DibPicture      =   "frmEstoque_Retirar_listaRE.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmEstoque_Retirar_listaRE.frx":62E4
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView Lista_RE 
      Height          =   2160
      Left            =   90
      TabIndex        =   2
      Top             =   450
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Data"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "n° RE"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Saldo"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmEstoque_Retirar_listaRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Load()
'On Error GoTo tratar_erro
'
'ProcVerificaTipoRequisicao
'
'
'
'Lista_RE.ListItems.Clear
'Set TBLISTA = CreateObject("adodb.recordset")
'TBLISTA.Open "Select CP.* FROM compras_pedido CP INNER JOIN compras_pedido_lista CPL ON CP.idpedido = CPL.idpedido where CPL.desenho = '" & frmestoque_Retirar.txtCodigo & "' and (CP.status_pedido = 'ABERTO' or CP.status_pedido = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
'If TBLISTA.EOF = False Then
'    Do While TBLISTA.EOF = False
'        With Lista.ListItems
'            .Add , , TBLISTA!IDpedido
'            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Pedido), "", TBLISTA!Pedido)
'            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Fornecedor), "", TBLISTA!Fornecedor)
'        End With
'        TBLISTA.MoveNext
'    Loop
'Else
'    USMsgBox ("Não foi encontrado nenhum pedido de compra com o status ABERTO ou PARCIAL para o produto " & frmestoque_Retirar.txtCodigo & "."), vbExclamation, "CAPRIND v5.0"
'End If
'TBLISTA.Close
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Private Sub ProcVerificaTipoRequisicao()
'On Error GoTo tratar_erro
'
''With frmestoque_Retirar
''If .Listamaterial.ListItems.Count = 0 Then Exit Sub
''.ProcLimpaCampos
''With .Listamaterial.SelectedItem
''    .txtCodigo.Text = .ListSubItems(6)
''    .txtQuant_prevista = .ListSubItems(1)
''    .txtQuant_prevista_PC = .ListSubItems(4)
''    .txtBaixado = .ListSubItems(2)
''    .txtRetirar = .ListSubItems(3)
''
''   ' If Cmb_RE <> "" Then txtquantretirado = txtQuant_prevista
''End With
'
''.ProcCarregaProduto
''If Expedir = True Then 'Retirar item de nota fiscal
''    ProcCarregaRE_NF
''    ProcCarregaLote_NF
''Else 'Retirar item de requisicao de materiais
''    If Requisicao_materiais = False Then ProcCarregaRE Else ProcCarregaRE_RM
''    If Requisicao_materiais = False Then ProcCarregaLote
''End If
''End If
''
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
