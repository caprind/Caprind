VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_centro_de_custo_consolidar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Custos - Centro de custo - Consolidar"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   315
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
   ScaleHeight     =   5430
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1395
      Left            =   55
      TabIndex        =   5
      Top             =   990
      Width           =   7455
      Begin VB.TextBox txtResponsavel 
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
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   375
         Width           =   5895
      End
      Begin VB.TextBox txtData 
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
         Top             =   375
         Width           =   1185
      End
      Begin VB.ComboBox Cmb_CC 
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
         Height          =   330
         ItemData        =   "Frm_centro_de_custo_consolidar.frx":0000
         Left            =   180
         List            =   "Frm_centro_de_custo_consolidar.frx":0019
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Centro de custo."
         Top             =   930
         Width           =   7095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3870
         TabIndex        =   10
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   9
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3150
         TabIndex        =   8
         Top             =   730
         Width           =   1155
      End
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
      Left            =   3785
      TabIndex        =   4
      Text            =   "0"
      Top             =   3870
      Visible         =   0   'False
      Width           =   1215
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "Frm_centro_de_custo_consolidar.frx":0047
      Count           =   1
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2745
      Left            =   60
      TabIndex        =   3
      Top             =   2400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4842
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Centro de custo"
         Object.Width           =   9322
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   7
      Top             =   5160
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   6
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
End
Attribute VB_Name = "Frm_centro_de_custo_consolidar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_consolidar As Boolean 'OK

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Usuarios_Setor.Codigo, Usuarios_Setor.Setor, Usuarios_Setor_Consolidacao.* from Usuarios_Setor_Consolidacao INNER JOIN Usuarios_Setor ON Usuarios_Setor_Consolidacao.ID_CC_consolidado = Usuarios_Setor.ID where Usuarios_Setor_Consolidacao.ID_CC = " & Frm_centro_de_custo.txtId & " order by Usuarios_Setor.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
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
                If USMsgBox("Deseja realmente excluir este(s) centro(s) de custo?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            
            Conexao.Execute "DELETE from Usuarios_Setor_Consolidacao where ID = " & .ListItems(InitFor)
            '====================================
            Modulo = "Custos/Centro de custo/Consolidar"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Centro de custo: " & .ListItems(InitFor).SubItems(1)
            Documento1 = ""
            ProcGravaEvento
            '===================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centos(s) de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Centro(s) de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame1.Enabled = False
    ProcCarregaLista
    Novo_consolidar = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_consolidar = True
Frame1.Enabled = True
Cmb_CC.SetFocus

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
Cmb_CC.ListIndex = -1
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
If Cmb_CC = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_CC.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Usuarios_Setor_Consolidacao where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_CC = Frm_centro_de_custo.txtId
If txtData <> "" Then TBGravar!Data = txtData Else TBGravar!Data = Date
If txtResponsavel <> "" Then TBGravar!Responsavel = txtResponsavel Else TBGravar!Responsavel = pubUsuario
TBGravar!ID_CC_consolidado = Cmb_CC.ItemData(Cmb_CC.ListIndex)
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
ProcCarregaLista
If Novo_consolidar = True Then
    USMsgBox ("Novo centro de custo consolidado com sucesso."), vbInformation, "CAPRIND v5.0"
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
Modulo = "Custos/Centro de custo/Consolidar"
ID_documento = txtId
Documento = "Centro de custo: " & Cmb_CC
Documento1 = ""
ProcGravaEvento
'==================================
Novo_consolidar = False

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
ProcLimpaVariaveisPrincipais
'ProcCarregaComboSetor Cmb_CC, "Setor is not null and ID <> " & Frm_centro_de_custo.txtID, "", False, True, True,False
With Cmb_CC
    .Clear
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "Select * from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = False Then
        TextoFiltro = "Select * from Usuarios_setor where ID_empresa = " & Frm_centro_de_custo.Cmb_empresa.ItemData(Frm_centro_de_custo.Cmb_empresa.ListIndex) & " and ID <> " & Frm_centro_de_custo.txtId & " order by Codigo"
    Else
        TextoFiltro = "Select US.* from Usuarios_setor US INNER JOIN Usuarios_Setor_Responsavel USR ON US.ID = USR.ID_CC where USR.Responsavel_CC = '" & pubUsuario & "' and US.ID_empresa = " & Frm_centro_de_custo.Cmb_empresa.ItemData(Frm_centro_de_custo.Cmb_empresa.ListIndex) & " and ID <> " & Frm_centro_de_custo.txtId & " order by US.Codigo"
    End If
    TBAcessos.Close
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If IsNull(TBAbrir!CODIGO) = False And TBAbrir!CODIGO <> "" Then
                .AddItem TBAbrir!Setor & " - " & TBAbrir!CODIGO
            Else
                .AddItem TBAbrir!Setor
            End If
            .ItemData(.NewIndex) = TBAbrir!ID
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
End With
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_consolidar = True Then
    If USMsgBox("O centro de custo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_consolidar = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_consolidar = False
Unload Me

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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Usuarios_Setor.Codigo, Usuarios_Setor.Setor, Usuarios_Setor_Consolidacao.* from Usuarios_Setor INNER JOIN Usuarios_Setor_Consolidacao ON Usuarios_Setor.ID = Usuarios_Setor_Consolidacao.ID_CC_consolidado where Usuarios_Setor_Consolidacao.ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
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

txtData = Format(TBLISTA!Data, "dd/mm/yy")
txtResponsavel = TBLISTA!Responsavel
txtId = TBLISTA!ID
If IsNull(TBLISTA!CODIGO) = False And TBLISTA!CODIGO <> "" Then Cmb_CC = TBLISTA!Setor & " - " & TBLISTA!CODIGO Else Cmb_CC = TBLISTA!Setor
Novo_consolidar = False
Frame1.Enabled = True

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
