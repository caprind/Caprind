VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_Localizar_Empresa 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Faturamento - Nota fiscal - Localizar empresa"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   ClipControls    =   0   'False
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
      Left            =   55
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
         Text            =   "Razão social"
         Object.Width           =   5868
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Endereço"
         Object.Width           =   5868
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Cidade"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "UF"
         Object.Width           =   970
      EndProperty
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_Localizar_Empresa"
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

If Clientes = True Then
    Caption = "Administrativo - Vendas - Cliente - Localizar empresa"
ElseIf Compras_Fornecedores = True Then
        Caption = "Administrativo - Compras - Fornecedores - Localizar empresa"
    ElseIf Compras_Pedido = True Then
            Caption = "Administrativo - Compras - Pedido - Localizar empresa"
        ElseIf Vendas_Proposta = True Then
                Caption = "Administrativo - Vendas - Proposta comercial - Localizar empresa"
            ElseIf Vendas_PI = True Then
                    Caption = "Administrativo - Vendas - Pedido interno - Localizar empresa"
                Else
                    If Faturamento = True Then
                        If Sit_REG = 4 Then
                            Caption = "Administrativo - Faturamento - Minuta de despacho - Localizar empresa"
                        Else
                            If Formulario = "Faturamento/Nota fiscal/Própria" Then
                                Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Localizar empresa"
                            ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
                                    Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Localizar empresa"
                                ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                                        Caption = "Estoque - Ordem de faturamento - Localizar empresa"
                                    Else
                                        Caption = "Estoque - Nota fiscal - Localizar empresa"
                            End If
                        End If
                    End If
End If
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Empresa order by Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Empresa), "", TBLISTA!Empresa)
            If IsNull(TBLISTA!Tipo_endereco) = False And TBLISTA!Tipo_endereco <> "" Then
                Endereco = TBLISTA!Tipo_endereco & ": " & IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
            Else
                Endereco = IIf(IsNull(TBLISTA!Endereco), "", TBLISTA!Endereco)
            End If
            .Item(.Count).SubItems(2) = Endereco
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Cidade), "", TBLISTA!Cidade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!UF), "", TBLISTA!UF)
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * FROM Empresa WHERE Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If Faturamento = True Then
        If Sit_REG < 4 Then
            With frmFaturamento_Prod_Serv
                Select Case Sit_REG
                    Case 1:
                        IDCliente = TBFornecedor!CODIGO
                        .txt_Razao.Text = IIf(IsNull(TBFornecedor!Empresa), "", TBFornecedor!Empresa)
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        .txt_Endereco.Text = Endereco
                        .txtNumero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
                        If IsNull(TBFornecedor!Tipo_bairro) = False And TBFornecedor!Tipo_bairro <> "" Then
                            Bairro = TBFornecedor!Tipo_bairro & ": " & IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
                        Else
                            Bairro = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
                        End If
                        .txt_Bairro.Text = Bairro
                        .txttipocliente = "E"
                        .txt_IE.Text = IIf(IsNull(TBFornecedor!ie), "", TBFornecedor!ie)
                        .txt_CNPJ_CPF.Text = IIf(IsNull(TBFornecedor!CNPJ), "", TBFornecedor!CNPJ)
                        .Txt_CEP.Text = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
                        .txt_Municipio.Text = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                        .cbo_UF.Text = IIf(IsNull(TBFornecedor!UF), "", TBFornecedor!UF)
                        '.txt_FoneFAX.Text = IIf(IsNull(TBFornecedor!Telefone), "", TBFornecedor!Telefone)
                        
                        .txtIDCliente.Text = IDCliente
                    Case 2:
                        .txtidinttransp = Lista.SelectedItem
                        .TxtTransp_nome.Text = Lista.SelectedItem.ListSubItems(1)
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        .txtTransp_endereco.Text = Endereco
                        .txtTransp_numero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
                        .txtTransp_municipio = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                        .txtTransp_uf_Transportadora = IIf(IsNull(TBFornecedor!UF), "", TBFornecedor!UF)
                        If IsNull(TBFornecedor!CNPJ) = True Or TBFornecedor!CNPJ = "__.___.___/____-__" Or TBFornecedor!CNPJ = "" Then .txtTransp_cnpj = "" Else .txtTransp_cnpj = TBFornecedor!CNPJ
                        .txtTransp_IE = IIf(IsNull(TBFornecedor!ie), "", TBFornecedor!ie)
                    Case 3:
                        If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                            Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        Else
                            Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                        End If
                        Redespacho = "Nome: " & Lista.SelectedItem.ListSubItems(1) & " - Endereço: " & Endereco & " - Número: " & IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero) & " - Cidade: " & IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade) & " - UF: " & IIf(IsNull(TBFornecedor!UF), "", TBFornecedor!UF) & " - CNPJ: " & IIf(IsNull(TBFornecedor!CNPJ), "", TBFornecedor!CNPJ) & " - IE: " & IIf(IsNull(TBFornecedor!ie), "", TBFornecedor!ie)
                        If .txtDados_DadosAdicionais <> "" Then
                            .txtDados_DadosAdicionais = .txtDados_DadosAdicionais & " | REDESPACHO: " & Redespacho
                        Else
                            .txtDados_DadosAdicionais = Redespacho
                        End If
                End Select
            End With
        Else
            With frmMinuta
                .txtID_transp = Lista.SelectedItem
                .txtTranportadora = Lista.SelectedItem.ListSubItems(1)
                If IsNull(TBFornecedor!Tipo_endereco) = False And TBFornecedor!Tipo_endereco <> "" Then
                    Endereco = TBFornecedor!Tipo_endereco & ": " & IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                Else
                    Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
                End If
                .txtendereco = Endereco
                .txtCidade = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
                .cmbuf = IIf(IsNull(TBFornecedor!UF), "", TBFornecedor!UF)
                .txttelefone = IIf(IsNull(TBFornecedor!Telefone), "", TBFornecedor!Telefone)
                .txtFax = IIf(IsNull(TBFornecedor!Fax), "", TBFornecedor!Fax)
                If IsNull(TBFornecedor!CNPJ) = True Or TBFornecedor!CNPJ = "__.___.___/____-__" Or TBFornecedor!CNPJ = "" Then .txtCNPJ = "" Else .txtCNPJ = TBFornecedor!CNPJ
                .txtIE = IIf(IsNull(TBFornecedor!ie), "", TBFornecedor!ie)
            End With
        End If
        Unload Me
        Exit Sub
    End If
    If Compras_Pedido = True Then
        frmCompras_Pedido.cmbtransporte = Lista.SelectedItem.ListSubItems(1)
        Unload Me
        Exit Sub
    End If
    If Clientes = True Then
        frmVendas_cliente.cmbtransportadora = Lista.SelectedItem.ListSubItems(1)
        Unload Me
        Exit Sub
    End If
    If Compras_Fornecedores = True Then
        frmCompras_fornecedores.cmbtransportadora = Lista.SelectedItem.ListSubItems(1)
        Unload Me
        Exit Sub
    End If
    If Vendas_PI = True Or Vendas_Proposta = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            .cmbtransportadora = Lista.SelectedItem.ListSubItems(1)
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

