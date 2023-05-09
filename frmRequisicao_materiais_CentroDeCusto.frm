VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmRequisicao_materiais_CentroDeCusto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Requisição de materiais - Centro de custo"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Height          =   855
      Left            =   55
      TabIndex        =   1
      Top             =   990
      Width           =   4185
      Begin VB.ComboBox cmbCentroDeCusto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Centro de custo."
         Top             =   390
         Width           =   3825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo"
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
         Left            =   1515
         TabIndex        =   2
         Top             =   180
         Width           =   1155
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
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
      ButtonWidth1    =   38
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   42
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "6"
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
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "7"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "8"
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
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2220
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmRequisicao_materiais_CentroDeCusto.frx":0000
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmRequisicao_materiais_CentroDeCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Permitido1 = False
With frmRequisicao_materiais
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
        
            'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
            If Permitido1 = False Then
                Permitido = False
                Set TBTempo = CreateObject("adodb.recordset")
                TBTempo.Open "Select Codigo from Empresa where Codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBTempo.EOF = False Then
                    Formulario = "Estoque/Autorização de centro de custo sem previsão"
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select ID_PC from projproduto where desenho = '" & .Lista.ListItems.Item(InitFor).ListSubItems(1) & "' and ID_PC IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = True Then
                        If USMsgBox("Existe(m) produto(s) sem conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                    Else
                        Set TBCQ = CreateObject("adodb.recordset")
                        TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & cmbCentroDeCusto.ItemData(cmbCentroDeCusto.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCQ.EOF = True Then
                            If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                        Else
                            Permitido = True
                        End If
                        TBCQ.Close
                    End If
                    TBProduto.Close
                    If Permitido = False Then Exit Sub
                End If
                TBTempo.Close
            End If
        
            Set TBCQ = CreateObject("adodb.recordset")
            TBCQ.Open "Select Desenho from projproduto where Desenho = '" & .Lista.ListItems(InitFor).SubItems(1) & "' and ID_CC = " & cmbCentroDeCusto.ItemData(cmbCentroDeCusto.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBCQ.EOF = True Then Conexao.Execute "Update Requisicao_materiais_lista Set ID_CC = " & cmbCentroDeCusto.ItemData(cmbCentroDeCusto.ListIndex) & " where IDLista = " & .Lista.ListItems(InitFor)
            TBCQ.Close
            '==================================
            Modulo = "Estoque/Requisição de materiais"
            Evento = "Cadastrar centro de custo"
            ID_documento = .Lista.ListItems(InitFor)
            Documento = "Nº requisição: " & .txtrequisicao
            Documento1 = "Cód. interno: " & .Lista.ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor

    USMsgBox ("Centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Set TBAbrir = CreateObject("adodb.recordset")
    If .TXTIDLista <> "" Then
        TBAbrir.Open "Select * from Requisicao_materiais_lista where idlista = " & .TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .procCarregaDados_Itens
        End If
        TBAbrir.Close
    End If
    For InitFor = 1 To .Lista.ListItems.Count
        .Lista.ListItems.Item(InitFor).Checked = False
    Next InitFor
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4185, 5, True

With frmRequisicao_materiais
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select A.* from Acessos A INNER JOIN Usuarios U ON A.IDUsuario = U.IDUsuario where U.Usuario = '" & .txtResponsavel & "' and A.Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = False Then
        ProcCarregaComboSetor cmbCentroDeCusto, "ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and setor is not null and (Consolidacao = 'False' or Consolidacao is null)", "", False, False, False, "", True, False
    Else
        ProcCarregaComboSetor cmbCentroDeCusto, "US.ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), "", False, False, False, .txtResponsavel, True, False
    End If
    TBAcessos.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
