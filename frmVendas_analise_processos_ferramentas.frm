VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_analise_processos_ferramentas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outros - Análise crítica - Processos - Utensílios da fase"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
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
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_ID_acessorio 
      BackColor       =   &H80000014&
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
      Left            =   8370
      MouseIcon       =   "frmVendas_analise_processos_ferramentas.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Text            =   "0"
      ToolTipText     =   "idferramenta."
      Top             =   4140
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtid 
      BackColor       =   &H80000014&
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
      Left            =   7740
      TabIndex        =   19
      Text            =   "0"
      Top             =   4140
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   55
      TabIndex        =   10
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtrev 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Revisão."
         Top             =   390
         Width           =   475
      End
      Begin VB.TextBox txtfase 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10710
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Fase."
         Top             =   390
         Width           =   645
      End
      Begin VB.TextBox txtcodinterno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   1905
      End
      Begin VB.TextBox txtmaquina 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11370
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Posto de trabalho."
         Top             =   390
         Width           =   3645
      End
      Begin VB.TextBox txtdescricao 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2592
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   8105
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12555
         TabIndex        =   18
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10860
         TabIndex        =   17
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6292
         TabIndex        =   12
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   682
         TabIndex        =   11
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   55
      TabIndex        =   13
      Top             =   1830
      Width           =   15195
      Begin VB.TextBox Txt_codinterno 
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
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   2055
      End
      Begin VB.CommandButton cmdProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2265
         Picture         =   "frmVendas_analise_processos_ferramentas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Localizar ferramentas."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_descricao 
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
         Left            =   2670
         TabIndex        =   8
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   10740
      End
      Begin VB.TextBox Txt_quant 
         Alignment       =   1  'Right Justify
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
         Left            =   13410
         TabIndex        =   7
         ToolTipText     =   "Quantidade."
         Top             =   390
         Width           =   1572
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13806
         TabIndex        =   16
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   757
         TabIndex        =   15
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7702
         TabIndex        =   14
         Top             =   180
         Width           =   690
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   20
      Top             =   9750
      Width           =   15195
      _ExtentX        =   26802
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
      Height          =   7065
      Left            =   60
      TabIndex        =   9
      Top             =   2670
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12462
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   19059
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   21
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
         Left            =   11640
         Top             =   195
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_analise_processos_ferramentas.frx":040C
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmVendas_analise_processos_ferramentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Processo_Ferramenta_ACC As Boolean 'OK

Private Sub cmdProduto_Click()
On Error GoTo tratar_erro

Vendas_Analise = True
Permitido = False
frmFerramentasdafase_localizar.Show 1
If Permitido = True Then Txt_quant.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: cmdProduto_Click
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

ProcCarregaToolBar1 Me, 15195, 7, True
With frmVendas_analise
    txtCodinterno.Text = .txtCodInterno_processos_item
    txtRev.Text = .txtRev_desenho
    txtFase.Text = .txtFase
    txtmaquina.Text = .txtMaquina_processos
    txtdescricao.Text = .txtdescricao
End With
ProcAtualizaFerramentas
       
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
                If USMsgBox("Deseja realmente excluir este(s) utensílio(s) da fase?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Vendas_analise_setores where ID = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Outros/Análise crítica/Processos/Utensílios da fase"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            With frmVendas_analise
                Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise & " - Fase: " & .txtFase
            End With
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) utensílio(s) da fase antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Utensílio(s) da fase excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizaFerramentas
    ProcLimpaCampos
End If
      
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmVendas_analise
    If .FunVerifStatusAnalise("criar novo utensílio da fase", True) = False Then Exit Sub
    If .FunVerifValidSetorAnalise("criar novo utensílio da fase", .SSTab1.Tab, True) = False Then Exit Sub
End With
Frame4.Enabled = True
ProcLimpaCampos
Novo_Processo_Ferramenta_ACC = True
cmdProduto_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
Txt_ID_acessorio = 0
Txt_codinterno = ""
Txt_descricao = ""
Txt_quant = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_codinterno.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdProduto_Click
    Exit Sub
End If
Quant = IIf(Txt_quant = "", 0, Txt_quant)
If Quant <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    Txt_quant.SetFocus
    Exit Sub
End If
With frmVendas_analise
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Vendas_analise_setores where ID = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
    Else
        If .FunVerifStatusAnalise("alterar o utensílio da fase", True) = False Then Exit Sub
        If .FunVerifValidSetorAnalise("alterar o utensílio da fase", .SSTab1.Tab, True) = False Then Exit Sub
    End If
    TBGravar!IDAnalise = .txtId
    TBGravar!Fase = .lista_Processos.SelectedItem.ListSubItems(1)
    TBGravar!ID_acessorio = IIf(Txt_ID_acessorio = "", 0, Txt_ID_acessorio)
    TBGravar!Codinterno = Txt_codinterno
    TBGravar!Qtde = Txt_quant
    TBGravar!Texto = Txt_descricao
    TBGravar!Tipo = "M"
    TBGravar!Setor = "FERRAMENTAS"
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select Unidade, Unidade_com, Pcusto from projproduto where Desenho = '" & Txt_codinterno & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        TBGravar!Un = TBProduto!Unidade
        TBGravar!Unidade_com = TBProduto!Unidade_com
        TBGravar!VlrUnit = IIf(IsNull(TBProduto!PCusto), 0, TBProduto!PCusto)
    End If
    TBProduto.Close
    TBGravar!vlrTotal = Format(TBGravar!VlrUnit * TBGravar!Qtde, "###,##0.00")
    
    TBGravar.Update
    txtId = TBGravar!ID
    TBGravar.Close
    
    If Novo_Processo_Ferramenta_ACC = True Then
        USMsgBox ("Novo utensílio da fase cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
        ProcAtualizaFerramentas
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar"
        ProcAtualizaFerramentas
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
    '==================================
    Modulo = "Outros/Análise crítica/Processos/Utensílios da fase"
    ID_documento = txtId
    Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise & " - Fase: " & .txtFase
    Documento1 = "Cód. interno: " & Txt_codinterno
    ProcGravaEvento
    '==================================
    Novo_Processo_Ferramenta_ACC = False
End With

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
                If frmVendas_analise.FunVerifStatusAnalise("", False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                If frmVendas_analise.FunVerifValidSetorAnalise("", frmVendas_analise.SSTab1.Tab, False) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
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

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If frmVendas_analise.FunVerifStatusAnalise("excluir este utensílio da fase", True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If frmVendas_analise.FunVerifValidSetorAnalise("excluir este utensílio da fase", frmVendas_analise.SSTab1.Tab, True) = False Then .ListItems.Item(InitFor).Checked = False
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
Set TBFerramentas = CreateObject("adodb.recordset")
TBFerramentas.Open "Select ID, ID_acessorio, Codinterno, Texto, Qtde from Vendas_analise_setores where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFerramentas.EOF = False Then
    ProcLimpaCampos
    txtId = TBFerramentas!ID
    Txt_ID_acessorio = IIf(IsNull(TBFerramentas!ID_acessorio), 0, TBFerramentas!ID_acessorio)
    Txt_codinterno.Text = IIf(IsNull(TBFerramentas!Codinterno), 0, TBFerramentas!Codinterno)
    Txt_descricao.Text = IIf(IsNull(TBFerramentas!Texto), "", TBFerramentas!Texto)
    Txt_quant.Text = IIf(IsNull(TBFerramentas!Qtde), "", Format(TBFerramentas!Qtde, "###,##0.0000"))
    CodigoLista = Lista.SelectedItem.index
    Frame4.Enabled = True
    Novo_Processo_Ferramenta_ACC = False
End If
TBFerramentas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaFerramentas()
On Error GoTo tratar_erro
  
Lista.ListItems.Clear
Set TBFerramentas = CreateObject("adodb.recordset")
TBFerramentas.Open "Select * from Vendas_analise_setores where IDanalise = " & frmVendas_analise.txtId & " and Fase = " & frmVendas_analise.lista_Processos.SelectedItem.ListSubItems(1) & " and Setor = 'FERRAMENTAS'", Conexao, adOpenKeyset, adLockOptimistic
If TBFerramentas.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBFerramentas.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBFerramentas.EOF = False
        With Lista.ListItems
            .Add , , TBFerramentas!ID
            .Item(.Count).SubItems(1) = TBFerramentas!Codinterno
            .Item(.Count).SubItems(2) = IIf(IsNull(TBFerramentas!Texto), "", TBFerramentas!Texto)
            .Item(.Count).SubItems(3) = Format(TBFerramentas!Qtde, "###,##0.00")
            If IsNull(TBFerramentas!ID_acessorio) = False And TBFerramentas!ID_acessorio <> 0 Then .Item(.Count).SubItems(4) = "Acessório" Else .Item(.Count).SubItems(4) = "Ferramenta"
        End With
        TBFerramentas.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBFerramentas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quant_Change()
On Error GoTo tratar_erro

If Txt_quant.Text <> "" Then
    VerifNumero = Txt_quant.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_quant.Text = ""
        Txt_quant.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quant_LostFocus()
On Error GoTo tratar_erro

If Txt_quant <> "" Then Txt_quant.Text = Format(Txt_quant.Text, "###,##0.0000")

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

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Processo_Ferramenta_ACC = True Then
    If USMsgBox("O utensílio ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Processo_Ferramenta_ACC = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Processo_Ferramenta_ACC = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

