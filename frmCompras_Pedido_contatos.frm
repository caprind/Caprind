VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Pedido_contatos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de contatos por fornecedor"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6285
   ControlBox      =   0   'False
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
   ScaleHeight     =   6540
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5070
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_Pedido_contatos.frx":0000
      Count           =   1
   End
   Begin VB.TextBox txtIDForn 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      MaxLength       =   60
      MouseIcon       =   "frmCompras_Pedido_contatos.frx":3403
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.TextBox txtIDContato 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1410
      MaxLength       =   60
      MouseIcon       =   "frmCompras_Pedido_contatos.frx":370D
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   950
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3570
      Left            =   60
      TabIndex        =   4
      Top             =   2700
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   6297
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Nome"
         Object.Width           =   5313
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Departamento"
         Object.Width           =   4344
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   55
      TabIndex        =   6
      Top             =   990
      Width           =   6165
      Begin VB.TextBox txtnomecont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         MaxLength       =   255
         TabIndex        =   0
         ToolTipText     =   "Nome."
         Top             =   180
         Width           =   4965
      End
      Begin VB.TextBox txtdepacont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Departamento."
         Top             =   540
         Width           =   4965
      End
      Begin VB.TextBox txttelcont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Telefone."
         Top             =   1260
         Width           =   4965
      End
      Begin ControlesUteis.txt txtemailcont 
         Height          =   360
         Left            =   990
         TabIndex        =   2
         ToolTipText     =   "E-mail."
         Top             =   900
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         Tamanho         =   4965
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   255
         BackColor       =   14737632
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   390
         TabIndex        =   10
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   9
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dpto. :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   8
         Top             =   540
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1260
         Width           =   735
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   12
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   1720
      ButtonCount     =   7
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Salvar (F3)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   38
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   39
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   118
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   122
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   6270
      Width           =   6165
      _ExtentX        =   10874
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
End
Attribute VB_Name = "frmCompras_Pedido_contatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Compras_Contato As Boolean 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) contato(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            
            Permitido = True
            Conexao.Execute "DELETE from Contatos_fornecedor where Idcontato = " & .ListItems(InitFor)

            '==================================
            Modulo = Formulario
            Evento = "Excluir contato"
            ID_documento = .ListItems(InitFor)
            Set TBFornecedor = CreateObject("adodb.recordset")
            TBFornecedor.Open "Select Nome_Razao from Compras_fornecedores where IDCliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
            If TBFornecedor.EOF = False Then
                Documento = "Cliente: " & Trim(TBFornecedor!Nome_Razao)
            End If
            TBFornecedor.Close
            Documento1 = "Contato: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) contato(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Contato(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame2.Enabled = False
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtnomecont.Text = "" Then
    USMsgBox ("Informe o nome do contato antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtnomecont.SetFocus
    Exit Sub
End If
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from Contatos_fornecedor where idcontato = " & txtIDContato.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = True Then TBFornecedor.AddNew
TBFornecedor!IDFornecedor = txtIDforn
TBFornecedor!Nome = txtnomecont.Text
TBFornecedor!Departamento = txtdepacont.Text
TBFornecedor!ramal = txttelcont.Text
TBFornecedor!Email = IIf(txtemailcont.Text = "", Null, txtemailcont.Text)
TBFornecedor.Update
txtIDContato = TBFornecedor!idcontato
TBFornecedor.Close
ProcCarregaLista
If Novo_Compras_Contato = True Then
    USMsgBox ("Novo contato cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo contato"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar contato"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = Formulario
ID_documento = txtIDContato
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select Nome_Razao from Compras_fornecedores where IDCliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    Documento = "Cliente: " & Trim(TBFornecedor!Nome_Razao)
End If
TBFornecedor.Close
Documento1 = "Contato: " & txtnomecont
ProcGravaEvento
'==================================
Novo_Compras_Contato = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Compras_Contato = True
Frame2.Enabled = True
txtnomecont.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6165, 7, True
ProcLimpaVariaveisPrincipais

If Compras_Cotacao = True Then
    txtIDforn = frmcompras_reqcot.txtIDforn
ElseIf Financeiro_Contas_Pagar = True Then
        txtIDforn = frmContas_Pagar.txtIDFornec
    ElseIf Financeiro_Contas_Pagar = True Then
            txtIDforn = frmContas_Pagar.txtIDFornec
        ElseIf Financeiro_Contas_Pagas = True Then
                txtIDforn = frmContas_Pagas.txtIDFornec
            ElseIf Financeiro_Contas_Receber = True Then
                    txtIDforn = frmContas_Receber.txtIDcliente
                ElseIf Financeiro_Contas_Recebidas = True Then
                        txtIDforn = frmContas_recebidas.txtIDcliente
                    Else
                        txtIDforn = frmCompras_Pedido.txtIDfornecedor
End If

'Carrega o telefone principal
If Compras_Cotacao = True Or Compras_Pedido = True Or Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Telefones from Compras_fornecedores where IDCliente = " & txtIDforn & " and Telefones IS NOT NULL and Telefones <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Caption = Caption & " (Tel.: " & TBLISTA!Telefones & ")"
    End If
Else
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Tel01 from Clientes where IDCliente = " & txtIDforn & " and Tel01 IS NOT NULL and Tel01 <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Caption = Caption & " (Tel.: " & TBLISTA!Tel01 & ")"
    End If
End If

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Compras_Contato = True Then
    If USMsgBox("O contato ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Compras_Contato = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Compras_Contato = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Contatos_fornecedor where IDfornecedor = " & txtIDforn & " order by nome", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!idcontato
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Nome), "", TBLISTA!Nome)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
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

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtIDContato = 0
txtnomecont = ""
txtdepacont = ""
txtemailcont.Text = ""
txttelcont = ""
CodigoLista = 0

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

If Lista.ListItems.Count = 0 Or Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Or Financeiro_Contas_Receber = True Or Financeiro_Contas_Recebidas = True Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from Contatos_fornecedor where idcontato = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If Compras_Cotacao = False Then
        With frmCompras_Pedido
            .txtContato.Text = IIf(IsNull(TBFornecedor!Nome), "", TBFornecedor!Nome)
            .txtEmail.Text = IIf(IsNull(TBFornecedor!Email), "", TBFornecedor!Email)
            .txttelefone.Text = IIf(IsNull(TBFornecedor!ramal), "", TBFornecedor!ramal)
        End With
    Else
        With frmcompras_reqcot
            .txtcontatoforn.Text = IIf(IsNull(TBFornecedor!Nome), "", TBFornecedor!Nome)
            .txttelforn.Text = IIf(IsNull(TBFornecedor!ramal), "", TBFornecedor!ramal)
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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Contatos_fornecedor where idcontato = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    Frame2.Enabled = True
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtIDContato = TBLISTA!idcontato
txtnomecont = IIf(IsNull(TBLISTA!Nome), "", TBLISTA!Nome)
txtdepacont = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
txtemailcont.Text = IIf(IsNull(TBLISTA!Email), "", TBLISTA!Email)
txttelcont = IIf(IsNull(TBLISTA!ramal), "", TBLISTA!ramal)
Novo_Compras_Contato = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyReturn: Lista_DblClick
    'Case vbKeyF1: cmdAjuda_Click
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
