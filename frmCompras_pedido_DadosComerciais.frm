VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_pedido_DadosComerciais 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compras - Pedido de compra - Dados comerciais - Texto padrão"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7575
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
   ScaleHeight     =   6435
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_pedido_DadosComerciais.frx":0000
      Count           =   1
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Text            =   "0"
      Top             =   4470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
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
      Height          =   2205
      Left            =   55
      TabIndex        =   5
      Top             =   990
      Width           =   7455
      Begin VB.TextBox txtdata 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox TxtResponsavel 
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
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   5865
      End
      Begin VB.TextBox txtTexto 
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
         Height          =   1095
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "Texto padrão."
         Top             =   960
         Width           =   7065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   690
         TabIndex        =   11
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3855
         TabIndex        =   10
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Texto*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3457
         TabIndex        =   6
         Top             =   750
         Width           =   510
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2265
      Left            =   60
      TabIndex        =   1
      Top             =   3870
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Aplic"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Texto"
         Object.Width           =   5406
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   44
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
      ButtonLeft3     =   83
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
      ButtonLeft4     =   124
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   128
      ButtonTop5      =   2
      ButtonWidth5    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   171
      ButtonTop6      =   2
      ButtonWidth6    =   30
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   203
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   6150
      Width           =   7455
      _ExtentX        =   13150
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Texto para pesquisa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   55
      TabIndex        =   12
      Top             =   3180
      Width           =   7455
      Begin VB.TextBox Txt_texto_pesquisa 
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
         Height          =   315
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Texto para pesquisa."
         Top             =   240
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmCompras_pedido_DadosComerciais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_DadosComerciais As Boolean 'OK

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

If Txt_texto_pesquisa <> "" Then TextoFiltro = "Texto like '%" & Txt_texto_pesquisa & "%' and " Else TextoFiltro = ""
If Sit_REG = 3 Then TextoFiltroTipo = "" Else TextoFiltroTipo = "Tipo = '" & Tipo & "' and "
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_proposta_dadoscomerciais_padrao where " & TextoFiltro & TextoFiltroTipo & "(Aplic = " & Aplic & " Or Aplic IS NULL) order by Texto", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Aplic), "", TBLISTA!Aplic)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
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

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) texto(s) padrão(ões)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Conexao.Execute "DELETE from vendas_proposta_dadoscomerciais_padrao where id = " & .ListItems(InitFor)
            '====================================
            If Compras_Cotacao = True Then Modulo = "Administrativo/Compras/Cotação"
            If Compras_Pedido = True Then Modulo = "Administrativo/Compras/Pedido"
            If Vendas_Proposta = True Then Modulo = "Administrativo/Vendas/Proposta comercial"
            If Vendas_PI = True Then Modulo = "Administrativo/Vendas/Pedido interno"
            If Compras_Produtos = True Then Modulo = "Compras/Produtos e serviços"
            If Vendas_Produtos = True Then Modulo = "Vendas/Produtos e serviços"
            If Engenharia_Produtos = True Then Modulo = "Engenharia/Produtos e serviços"
            Evento = "Excluir cadastro de texto padrão"
            ID_documento = .ListItems(InitFor)
            Documento = "Texto: " & .ListItems(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '===================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) texto(s) padrão(ões) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Texto(s) padrão(ões) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Novo_DadosComerciais = False
    Frame1.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_DadosComerciais = True
Frame1.Enabled = True
txtTexto.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtTexto = ""
CodigoLista = 0
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If txtTexto.Text = "" Then
    USMsgBox ("Informe o texto padrão antes de salvar."), vbExclamation, "CAPRIND v5.0"
    txtTexto.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_proposta_Dadoscomerciais_padrao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Data = IIf(txtData = "", Date, txtData)
TBGravar!Responsavel = IIf(txtResponsavel = "", pubUsuario, txtResponsavel)
TBGravar!Aplic = Aplic
If Tipo <> "" Then TBGravar!Tipo = Tipo
TBGravar!Texto = txtTexto
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
ProcCarregaLista
If Novo_DadosComerciais = True Then
    USMsgBox ("Novo texto padrão cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo cadastro de texto padrão"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar cadastro de texto padrão"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
If Compras_Cotacao = True Then Modulo = "Administrativo/Compras/Cotação"
If Compras_Pedido = True Then Modulo = "Administrativo/Compras/Pedido"
If Vendas_Proposta = True Then Modulo = "Administrativo/Vendas/Proposta comercial"
If Compras_Fornecedores = True Then Modulo = "Administrativo/Compras/Fornecedores"
If Vendas_PI = True Then Modulo = "Administrativo/Vendas/Pedido interno"
If Sit_REG = 1 Then Modulo = "Compras/Produtos e serviços"
If Sit_REG = 2 Then Modulo = "Vendas/Produtos e serviços"
If Sit_REG = 3 Then Modulo = "Engenharia/Produtos e serviços"
ID_documento = txtId
Documento = "Texto: " & txtTexto
Documento1 = ""
ProcGravaEvento
'==================================
Novo_DadosComerciais = False

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
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7455, 7, True
If Compras_Cotacao = True Then
    Caption = "Compras - Cotação - Dados comerciais - Texto padrão"
    Tipo = "C"
ElseIf Compras_Fornecedores = True Then
    Caption = "Compras - Fornecedores - Dados comerciais - Texto padrão"
    Tipo = "C"
ElseIf Compras_Pedido = True Then
    Caption = "Compras - Pedido - Dados comerciais - Texto padrão"
    Tipo = "C"
ElseIf Vendas_Proposta = True Then
    Caption = "Vendas - Proposta comercial - Texto padrão"
    Tipo = "V"
ElseIf Vendas_PI = True Then
    Caption = "Vendas - Pedido interno - Texto padrão"
    Tipo = "V"
ElseIf Estoque_recebimento = True Then
    Caption = "Estoque - Recebimento - Dados comerciais - Texto padrão"
    Tipo = "C"
ElseIf Clientes = True Then
    Caption = "Vendas - Clientes - Dados comerciais - Texto padrão"
    Tipo = "V"
ElseIf Sit_REG = 1 Then
    Caption = "Compras - Produtos e serviços - Texto padrão"
    Tipo = "C"
ElseIf Sit_REG = 2 Then
    Caption = "Vendas - Produtos e serviços - Texto padrão"
    Tipo = "V"
ElseIf Sit_REG = 3 Then
    Caption = "Engenharia - Produtos e serviços - Texto padrão"
    Tipo = ""
End If
ProcLimpaVariaveisPrincipais
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_DadosComerciais = True Then
    If USMsgBox("O texto padrão ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_DadosComerciais = True Then Exit Sub Else Unload Me
    End If
End If
Novo_DadosComerciais = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from vendas_proposta_dadoscomerciais_padrao where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    If Compras_Cotacao = True Then
        Select Case TBLISTA!Aplic
            Case 1: frmcompras_reqcot.txtcondpagtoforn = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
        End Select
    ElseIf Compras_Pedido = True Then
        With frmCompras_Pedido
            Select Case TBLISTA!Aplic
                Case 1: .cmbpagamento = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 2: .txtprazo = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 3: .txtembalagem = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 4: .txtEscopo = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 5: .txttransporte = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            End Select
        End With
    ElseIf Vendas_PI = True Or Vendas_Proposta = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            Select Case TBLISTA!Aplic
                Case 1: .txtCondicoes = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 2: .txtprazo = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 3: .txtcalculos = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 4: .txtinspecao = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 5: .txtembalagem = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 6: .txttransporte = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 7: .txtimpostos = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 8: .txtgarantia = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 9: .txtReajuste = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 10: .txtValidade = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 11: .txtEscopo = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 12: .txtGravacao = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            End Select
        End With
    ElseIf Estoque_recebimento = True Then
        Select Case TBLISTA!Aplic
            Case 1: frmEstoque_Recebimento.txtcondpagamento = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
        End Select
    ElseIf Clientes = True Then
        With frmVendas_cliente
            Select Case TBLISTA!Aplic
                Case 1: .txtCondicoes = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 2: .txtprazo = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 3: .txtcalculos = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 4: .txtinspecao = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 5: .txtembalagem = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 6: .txttransporte = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 7: .txtimpostos = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 8: .txtgarantia = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 9: .txtReajuste = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 10: .txtValidade = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            End Select
        End With
    ElseIf Compras_Fornecedores = True Then
        With frmCompras_fornecedores
            Select Case TBLISTA!Aplic
                Case 1: .txtCondicoes = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 2: .txtprazo = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
               ' Case 3: .txtcalculos = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 4: .txtinspecao = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 5: .txtembalagem = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 6: .txttransporte = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 7: .txtimpostos = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
               ' Case 8: .txtgarantia = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                'Case 9: .txtReajuste = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 10: .txtValidade = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            End Select
        End With

    ElseIf Sit_REG = 1 Or Sit_REG = 2 Or Sit_REG = 3 Then
        With frmproj_produto
            Select Case TBLISTA!Aplic
                Case 4: .txtinspecao = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 5: .txtembalagem = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
                Case 12: .txtGravacao = IIf(IsNull(TBLISTA!Texto), "", TBLISTA!Texto)
            End Select
        End With
    End If
End If
TBLISTA.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from vendas_proposta_Dadoscomerciais_padrao where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId = TBAbrir!ID
txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtTexto = IIf(IsNull(TBAbrir!Texto), "", TBAbrir!Texto)
Frame1.Enabled = True
Novo_DadosComerciais = False

If Engenharia_Produtos = True Then Exit Sub

If IsNull(TBAbrir!Tipo) = True Or TBAbrir!Tipo = "" Then
    If Tipo = "C" Then
        If USMsgBox("Este texto padrão se aplica para o módulo de compras e vendas?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Conexao.Execute "Update vendas_proposta_Dadoscomerciais_padrao Set Tipo = 'C' where Id = " & txtId
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from vendas_proposta_Dadoscomerciais_padrao where Texto = '" & txtTexto & "' and Tipo = 'V'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then TBGravar.AddNew
            TBGravar!Aplic = Aplic
            TBGravar!Texto = txtTexto
            TBGravar!Tipo = "V"
            TBGravar.Update
            TBGravar.Close
        ElseIf USMsgBox("Este texto padrão se aplica para o módulo de compras?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                Conexao.Execute "Update vendas_proposta_Dadoscomerciais_padrao Set Tipo = 'C' where Id = " & txtId
            Else
                Conexao.Execute "Update vendas_proposta_Dadoscomerciais_padrao Set Tipo = 'V' where Id = " & txtId
        End If
    Else
        If USMsgBox("Este texto padrão se aplica para o módulo de vendas e compras?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Conexao.Execute "Update vendas_proposta_Dadoscomerciais_padrao Set Tipo = 'V' where Id = " & txtId
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from vendas_proposta_Dadoscomerciais_padrao where Texto = '" & txtTexto & "' and Tipo = 'C'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then TBGravar.AddNew
            TBGravar!Aplic = Aplic
            TBGravar!Texto = txtTexto
            TBGravar!Tipo = "C"
            TBGravar.Update
            TBGravar.Close
        ElseIf USMsgBox("Este texto padrão se aplica para o módulo de vendas?", vbYesNo, "CAPRIND v5.0") = vbYes Then
                Conexao.Execute "Update vendas_proposta_Dadoscomerciais_padrao Set Tipo = 'V' where Id = " & txtId
            Else
                Conexao.Execute "Update vendas_proposta_Dadoscomerciais_padrao Set Tipo = 'C' where Id = " & txtId
        End If
    End If
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_texto_pesquisa_Change()
On Error GoTo tratar_erro

ProcCarregaLista

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
