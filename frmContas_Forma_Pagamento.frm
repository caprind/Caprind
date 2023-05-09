VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContas_Forma_Pagamento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Contas a pagar - Forma da baixa"
   ClientHeight    =   5505
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5430
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmContas_Forma_Pagamento.frx":0000
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
      TabIndex        =   6
      Text            =   "0"
      Top             =   3870
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
      Height          =   1425
      Left            =   55
      TabIndex        =   4
      Top             =   990
      Width           =   7455
      Begin VB.TextBox txtdata 
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   1215
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   5835
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
         Height          =   315
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Texto padrão."
         Top             =   990
         Width           =   7065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   615
         TabIndex        =   8
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
         Index           =   11
         Left            =   3870
         TabIndex        =   7
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
         TabIndex        =   5
         Top             =   780
         Width           =   510
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2775
      Left            =   55
      TabIndex        =   3
      Top             =   2430
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4895
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Texto"
         Object.Width           =   11933
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   9
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   36
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
      ButtonLeft2     =   40
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   86
      ButtonTop3      =   2
      ButtonWidth3    =   45
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
      ButtonLeft4     =   133
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
      ButtonLeft5     =   137
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
      ButtonLeft6     =   180
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
      ButtonLeft7     =   212
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   10
      Top             =   5220
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
End
Attribute VB_Name = "frmContas_Forma_Pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Forma_Pgto As Boolean 'OK

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_FormaPagto where Tipo = '" & Tipo & "' order by Descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!IdForma
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
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

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) texto(s) padrão(ões)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from tbl_FormaPagto where IdForma = " & .ListItems(InitFor)
            '====================================
            Modulo = Formulario
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Texto: " & .ListItems(InitFor).SubItems(1)
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
    Frame1.Enabled = False
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Novo_Forma_Pgto = False Then
    ProcLimpaCampos
    Novo_Forma_Pgto = True
End If
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

If Financeiro_Contas_Pagar = True Then
    Caption = "Financeiro - Contas a pagar - Forma da baixa"
ElseIf Financeiro_Forma_Pgto_Pagar = True Then
        Caption = "Financeiro - Contas a pagar - Baixar - Forma da baixa"
    ElseIf Financeiro_Contas_Receber = True Then
            Caption = "Financeiro - Contas a receber - Forma da baixa"
        ElseIf Financeiro_Forma_Pgto_Receber = True Then
                Caption = "Financeiro - Contas a receber - Baixar - Forma da baixa"
            Else
                Caption = "Financeiro - Instituições - Forma da baixa"
End If

    
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
    USMsgBox ("Informe o texto padrão antes de salvar."), vbInformation, "CAPRIND v5.0"
    txtTexto.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_FormaPagto where IdForma = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If TBGravar!Descricao = "CHEQUE" Or TBGravar!Descricao = "CHEQUE PRÉ-DATADO" Or TBGravar!Descricao = "DOC" Or TBGravar!Descricao = "TED" Or TBGravar!Descricao = "SAQUE" Then
        USMsgBox ("Não é permitido alterar este texto padrão."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    If TBGravar!Descricao <> txtTexto Then
        Conexao.Execute "Update tbl_ContasPagar Set FormaBaixa = '" & txtTexto & "' where FormaBaixa = '" & TBGravar!Descricao & "'"
        Conexao.Execute "Update tbl_contas_receber Set FormaBaixa = '" & txtTexto & "' where FormaBaixa = '" & TBGravar!Descricao & "'"
        Conexao.Execute "Update tbl_Detalhes_Recebimento Set txt_tipoPagto = '" & txtTexto & "' where txt_tipoPagto = '" & TBGravar!Descricao & "'"
    End If
Else
    TBGravar.AddNew
End If
If txtData = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData
If txtResponsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel
TBGravar!Tipo = Tipo
TBGravar!Descricao = txtTexto.Text
TBGravar.Update
txtId = TBGravar!IdForma
TBGravar.Close
ProcCarregaLista
If Novo_Forma_Pgto = True Then
    USMsgBox ("Novo texto padrão cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = Formulario
ID_documento = txtId
Documento = "Texto: " & txtTexto
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Forma_Pgto = False

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
    'Case vbKeyF1: Ajuda
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
If Financeiro_Contas_Pagar = True Then
    Caption = "Financeiro - Contas a pagar - Forma da baixa"
    Formulario = "Financeiro/Contas a pagar/Forma da baixa"
    Tipo = "P"
ElseIf Financeiro_Forma_Pgto_Pagar = True Then
        Caption = "Financeiro - Contas a pagar - baixar - Forma da baixa"
        Formulario = "Financeiro/Contas a pagar/Baixar/Forma da baixa"
        Tipo = "P"
    ElseIf Financeiro_Contas_Receber = True Then
            Caption = "Financeiro - Contas a receber - Forma da baixa"
            Formulario = "Financeiro/Contas a receber/Forma da baixa"
            Tipo = "R"
        ElseIf Financeiro_Forma_Pgto_Receber = True Then
                Caption = "Financeiro - Contas a receber - Baixar - Forma da baixa"
                Formulario = "Financeiro/Contas a receber/Baixar/Forma da baixa"
                Tipo = "R"
            Else
                Caption = "Financeiro - Instituições - Forma da baixa"
                Formulario = "Financeiro/Instituições/Forma da baixa"
                If frm_Instituicoes.Cmb_operacao = "Débito" Then Tipo = "P" Else Tipo = "R"
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

If Novo_Forma_Pgto = True Then
    If USMsgBox("O texto padrão ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Forma_Pgto = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Forma_Pgto = False
If Financeiro_Contas_Pagar = True Then
    frmContas_Pagar.ProcCarregaComboForma
ElseIf Financeiro_Forma_Pgto_Pagar = True Then
        frm_Baixas.ProcCarregaComboForma
    ElseIf Financeiro_Contas_Receber = True Then
            frmContas_Receber.ProcCarregaComboForma
        ElseIf Financeiro_Forma_Pgto_Receber = True Then
                frm_Baixas_Receber.ProcCarregaComboForma
            Else
                frm_Instituicoes.ProcCarregaComboForma
End If
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
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_FormaPagto where IdForma = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!Descricao = "CHEQUE" Or TBAbrir!Descricao = "CHEQUE PRÉ-DATADO" Or TBAbrir!Descricao = "DOC" Or TBAbrir!Descricao = "TED" Or TBAbrir!Descricao = "SAQUE" Then
                        TBAbrir.Close
                        GoTo Proximo
                    End If
                End If
                TBAbrir.Close
                ProcVerificaRegistroUtilizadoSemMsg "tbl_ContasPagar", "FormaBaixa = '" & .ListItems(InitFor).SubItems(1) & "'"
                If Permitido = False Then GoTo Proximo
                ProcVerificaRegistroUtilizadoSemMsg "tbl_contas_receber", "FormaBaixa = '" & .ListItems(InitFor).SubItems(1) & "'"
                If Permitido = False Then GoTo Proximo
                ProcVerificaRegistroUtilizadoSemMsg "tbl_Detalhes_Recebimento", "txt_tipoPagto = '" & .ListItems(InitFor).SubItems(1) & "'"
                If Permitido = False Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
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
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from tbl_FormaPagto where IdForma = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If Financeiro_Contas_Pagar = True Then
        frmContas_Pagar.ProcCarregaComboForma
        If IsNull(TBFIltro!Descricao) = False And TBFIltro!Descricao <> "" Then frmContas_Pagar.cmb_forma = TBFIltro!Descricao
    ElseIf Financeiro_Forma_Pgto_Pagar = True Then
            frm_Baixas.ProcCarregaComboForma
            If IsNull(TBFIltro!Descricao) = False And TBFIltro!Descricao <> "" Then frm_Baixas.cmb_forma = TBFIltro!Descricao
        ElseIf Financeiro_Contas_Receber = True Then
                frmContas_Receber.ProcCarregaComboForma
                If IsNull(TBFIltro!Descricao) = False And TBFIltro!Descricao <> "" Then frmContas_Receber.cmb_forma = TBFIltro!Descricao
            ElseIf Financeiro_Forma_Pgto_Receber = True Then
                    frm_Baixas_Receber.ProcCarregaComboForma
                    If IsNull(TBFIltro!Descricao) = False And TBFIltro!Descricao <> "" Then frm_Baixas_Receber.cmb_forma = TBFIltro!Descricao
                Else
                    frm_Instituicoes.ProcCarregaComboForma
                    If IsNull(TBFIltro!Descricao) = False And TBFIltro!Descricao <> "" Then frm_Instituicoes.cmb_forma1 = TBFIltro!Descricao
    End If
End If
TBFIltro.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_FormaPagto where IdForma = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!Descricao = "CHEQUE" Or TBAbrir!Descricao = "CHEQUE PRÉ-DATADO" Or TBAbrir!Descricao = "DOC" Or TBAbrir!Descricao = "TED" Or TBAbrir!Descricao = "SAQUE" Then
                    USMsgBox ("Não é permitido excluir este texto padrão, pois o mesmo é padrão do sistema."), vbExclamation, "CAPRIND v5.0"
                    TBAbrir.Close
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            End If
            Mensagem = "Não é permitido excluir este texto padrão, pois o mesmo está sendo utilizado no módulo"
            ProcVerificaRegistroUtilizado "tbl_ContasPagar", "FormaBaixa = '" & .ListItems(InitFor).SubItems(1) & "'", "Financeiro/Contas a pagar"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_contas_receber", "FormaBaixa = '" & .ListItems(InitFor).SubItems(1) & "'", "Financeiro/Contas a receber"
            If Permitido = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            ProcVerificaRegistroUtilizado "tbl_Detalhes_Recebimento", "txt_tipoPagto = '" & .ListItems(InitFor).SubItems(1) & "'", "Faturamento/Nota fiscal"
            If Permitido = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_FormaPagto where IdForma = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
Frame1.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId = TBLISTA!IdForma
txtData = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
txtTexto = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
Novo_Forma_Pgto = False

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
