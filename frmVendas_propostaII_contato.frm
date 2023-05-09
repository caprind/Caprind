VERSION 5.00
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_propostaII_contato 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   " Lista de contatos"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   688
      DibPicture      =   "frmVendas_propostaII_contato.frx":0000
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
      Icon            =   "frmVendas_propostaII_contato.frx":37515
   End
   Begin VB.TextBox txtIDCliente 
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
      MouseIcon       =   "frmVendas_propostaII_contato.frx":3782F
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
      MouseIcon       =   "frmVendas_propostaII_contato.frx":37B39
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   950
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3915
      Left            =   90
      TabIndex        =   4
      Top             =   3180
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   6906
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
         Object.Width           =   512
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   90
      TabIndex        =   12
      Top             =   7140
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   90
      TabIndex        =   13
      Top             =   510
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4710
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_propostaII_contato.frx":37E43
         Count           =   1
      End
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
      Height          =   1665
      Left            =   90
      TabIndex        =   6
      Top             =   1500
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
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1260
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmVendas_propostaII_contato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Vendas_Contato As Boolean 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) contato(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            
            Permitido = True
            
            Conexao.Execute "DELETE from Clientes_Contatos where Idcontato = " & .ListItems(InitFor)
            '==================================
            Modulo = Formulario
            Evento = "Excluir contato"
            ID_documento = .ListItems(InitFor)
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select NomeRazao from Clientes where IDCliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
            If TBClientes.EOF = False Then
                Documento = "Cliente: " & Trim(TBClientes!NomeRazao)
            End If
            TBClientes.Close
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
    USMsgBox ("Informe o nome do contato antes de salvar."), vbInformation, "CAPRIND v5.0"
    txtnomecont.SetFocus
    Exit Sub
End If
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from clientes_contatos where idcontato = " & txtIDContato.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = True Then
    TBClientes.AddNew
Else
    If txtnomecont <> TBClientes!NomeContato Then
        Conexao.Execute "Update Vendas_tele Set Contato = '" & txtnomecont & "' where IDcliente = " & txtIDcliente & " and Contato = '" & TBClientes!NomeContato & "'"
        Conexao.Execute "Update Vendas_tele Set contato_visita = '" & txtnomecont & "' where IDcliente = " & txtIDcliente & " and contato_visita = '" & TBClientes!NomeContato & "'"
    End If
End If
TBClientes!IDCliente = txtIDcliente
TBClientes!NomeContato = txtnomecont.Text
TBClientes!Departamento = txtdepacont.Text
TBClientes!telefone = txttelcont.Text
TBClientes!Email = IIf(txtemailcont.Text = "", Null, txtemailcont.Text)
TBClientes.Update
txtIDContato = TBClientes!idcontato
TBClientes.Close
ProcCarregaLista
If Novo_Vendas_Contato = True Then
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
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select NomeRazao from Clientes where IDCliente = " & txtIDcliente, Conexao, adOpenKeyset, adLockOptimistic
If TBClientes.EOF = False Then
    Documento = "Cliente: " & Trim(TBClientes!NomeRazao)
End If
TBClientes.Close
Documento1 = "Contato: " & txtnomecont
ProcGravaEvento
'==================================
Novo_Vendas_Contato = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Vendas_Contato = True
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

If Vendas_Proposta = True Then
    txtIDcliente = frmVendas_proposta.txtIDcliente
ElseIf Vendas_PI = True Then
        txtIDcliente = frmVendas_PI.txtIDcliente
    ElseIf Telemarketing = True Then
            txtIDcliente = frmVendas_Tele_Clientes.txtIDcliente
        ElseIf Qualidade_PPAP_PSW = True Then
                txtIDcliente = frmQualidadePPAP.txtIDcliente
            ElseIf Analise_critica = True Then
                    txtIDcliente = frmVendas_analise.txtIDcliente
                ElseIf Financeiro_Contas_Pagar = True Then
                        txtIDcliente = frmContas_Pagar.txtIDFornec
                    ElseIf Financeiro_Contas_Pagas = True Then
                            txtIDcliente = frmContas_Pagas.txtIDFornec
                        ElseIf Financeiro_Contas_Receber = True Then
                                txtIDcliente = frmContas_Receber.txtIDcliente
                            Else
                                txtIDcliente = frmContas_recebidas.txtIDcliente
End If

'Carrega o telefone principal
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Telefones from Compras_fornecedores where IDCliente = " & txtIDcliente & " and Telefones IS NOT NULL and Telefones <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        USForm1.Caption = USForm1.Caption & " (Tel.:" & TBLISTA!Telefones & ")"
    End If
Else
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select Tel01 from Clientes where IDCliente = " & txtIDcliente & " and Tel01 IS NOT NULL and Tel01 <> N''", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        USForm1.Caption = USForm1.Caption & " (Tel.:" & TBLISTA!Tel01 & ")"
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

If Novo_Vendas_Contato = True Then
    If USMsgBox("O contato ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Vendas_Contato = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Vendas_Contato = False
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
TBLISTA.Open "Select * from clientes_contatos where IDCliente = " & txtIDcliente & " order by nomecontato", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!idcontato
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!NomeContato), "", TBLISTA!NomeContato)
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

If Lista.ListItems.Count = 0 Or Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Or Financeiro_Contas_Receber = True Or Financeiro_Contas_Recebidas = True Then Exit Sub
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from clientes_contatos where idcontato = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If Telemarketing = True Then
        With frmVendas_Tele_Clientes
            If Sit_REG = 1 Then
                .Txt_contato = IIf(IsNull(TBFornecedor!NomeContato), "", TBFornecedor!NomeContato)
                .Txt_departamento_contato = IIf(IsNull(TBFornecedor!Departamento), "", TBFornecedor!Departamento)
                .Txt_telefonel_contato = IIf(IsNull(TBFornecedor!telefone), "", TBFornecedor!telefone)
                .Txt_email_contato = IIf(IsNull(TBFornecedor!Email), "", TBFornecedor!Email)
            Else
                .Txt_contato_visita.Text = IIf(IsNull(TBFornecedor!NomeContato), "", TBFornecedor!NomeContato)
                .Txt_departamento_contato_visita.Text = IIf(IsNull(TBFornecedor!Departamento), "", TBFornecedor!Departamento)
            End If
        End With
    End If
    If Vendas_Proposta = True Or Vendas_PI = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            .txtRemetente.Text = IIf(IsNull(TBFornecedor!NomeContato), "", TBFornecedor!NomeContato)
            .txtdepartamento.Text = IIf(IsNull(TBFornecedor!Departamento), "", TBFornecedor!Departamento)
            If IsNull(TBFornecedor!telefone) = False And TBFornecedor!telefone <> "" Then .txttelefone.Text = TBFornecedor!telefone
            If IsNull(TBFornecedor!Email) = False And TBFornecedor!Email <> "" Then .txtEmail.Text = TBFornecedor!Email
        End With
    End If
    If Analise_critica = True Then
        With frmVendas_analise
            .txtContato.Text = IIf(IsNull(TBFornecedor!NomeContato), "", TBFornecedor!NomeContato)
            .txtdepartamento.Text = IIf(IsNull(TBFornecedor!Departamento), "", TBFornecedor!Departamento)
            If IsNull(TBFornecedor!telefone) = False And TBFornecedor!telefone <> "" Then .txttelefone.Text = TBFornecedor!telefone
            If IsNull(TBFornecedor!Email) = False And TBFornecedor!Email <> "" Then .txtEmail.Text = TBFornecedor!Email
        End With
    End If
    If Qualidade_PPAP_PSW = True Then
        frmQualidadePPAP.txtContato = IIf(IsNull(TBFornecedor!NomeContato), "", TBFornecedor!NomeContato)
        frmQualidadePPAP.txtEmail = IIf(IsNull(TBFornecedor!Email), "", TBFornecedor!Email)
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
TBLISTA.Open "Select * from clientes_contatos where idcontato = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
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
txtnomecont = IIf(IsNull(TBLISTA!NomeContato), "", TBLISTA!NomeContato)
txtdepacont = IIf(IsNull(TBLISTA!Departamento), "", TBLISTA!Departamento)
txtemailcont.Text = IIf(IsNull(TBLISTA!Email), "", TBLISTA!Email)
txttelcont = IIf(IsNull(TBLISTA!telefone), "", TBLISTA!telefone)
Frame2.Enabled = True
Novo_Vendas_Contato = False

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
    'Case vbKeyF1: ProcAjuda
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
