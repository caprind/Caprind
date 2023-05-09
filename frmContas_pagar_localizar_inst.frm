VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmContas_pagar_localizar_inst 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Contas à pagar - Localizar instituição"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   ClipControls    =   0   'False
   Icon            =   "frmContas_pagar_localizar_inst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   0
      Top             =   4620
      Width           =   9795
      _ExtentX        =   17277
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
   Begin MSComctlLib.ListView Lista 
      Height          =   4545
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   8017
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
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Banco"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Agência"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Conta"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   9710
      EndProperty
   End
End
Attribute VB_Name = "frmContas_pagar_localizar_inst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: Lista_DblClick
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

TextoFiltro = " and DtValidacao IS NOT NULL and Bloqueado <> 'True'"
If Financeiro_Contas_Pagar = True Then
    ID_empresa = frmContas_Pagar.Cmb_empresa.ItemData(frmContas_Pagar.Cmb_empresa.ListIndex)
ElseIf Financeiro_Contas_Pagas = True Then
        Caption = "Administrativo - Financeiro - Contas pagas - Localizar instituição"
        ID_empresa = frmContas_Pagas.Cmb_empresa.ItemData(frmContas_Pagas.Cmb_empresa.ListIndex)
        TextoFiltro = ""
    ElseIf Financeiro_Contas_Receber = True Then
            Caption = "Administrativo - Financeiro - Contas a receber - Localizar instituição"
            ID_empresa = frmContas_Receber.Cmb_empresa.ItemData(frmContas_Receber.Cmb_empresa.ListIndex)
        ElseIf Financeiro_Contas_Recebidas = True Then
                Caption = "Administrativo - Financeiro - Contas recebidas - Localizar instituição"
                ID_empresa = frmContas_recebidas.Cmb_empresa.ItemData(frmContas_recebidas.Cmb_empresa.ListIndex)
                TextoFiltro = ""
End If
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_instituicoes where ID_empresa = " & ID_empresa & TextoFiltro & " order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    With Lista.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!int_NBanco), "", TBLISTA!int_NBanco)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!txt_Agencia), "", TBLISTA!txt_Agencia)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!txt_conta), "", TBLISTA!txt_conta)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        Loop
    End With
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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * FROM tbl_instituicoes WHERE ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If Financeiro_Contas_Pagar = True Then
        With frmContas_Pagar
            .Cmb_tipo = "Instituição bancária"
            .txtIDFornec = TBFornecedor!ID
        End With
    ElseIf Financeiro_Contas_Pagas = True Then
        With frmContas_Pagas
            .Cmb_tipo = "Instituição bancária"
            .txtIDFornec = TBFornecedor!ID
        End With
    ElseIf Financeiro_Contas_Receber = True Then
            With frmContas_Receber
                .Cmb_tipo = "Instituição bancária"
                .txtIDCliente = TBFornecedor!ID
            End With
        ElseIf Financeiro_Contas_Recebidas = True Then
                With frmContas_recebidas
                    .Cmb_tipo = "Instituição bancária"
                    .txtIDCliente = TBFornecedor!ID
                End With
    End If
End If
TBFornecedor.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
