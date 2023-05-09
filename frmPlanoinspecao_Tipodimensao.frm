VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanoinspecao_Tipodimensao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Plano de inspeção - Tipo da dimensão"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7575
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   900
      TabIndex        =   4
      Text            =   "0"
      Top             =   2610
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3495
      Left            =   55
      TabIndex        =   1
      Top             =   1830
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
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
         Text            =   "Tipo da dimensão"
         Object.Width           =   11968
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   825
      Left            =   55
      TabIndex        =   2
      Top             =   990
      Width           =   7455
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
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Tipo da dimensão."
         Top             =   390
         Width           =   7095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo da dimensão"
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
         Left            =   3105
         TabIndex        =   3
         Top             =   180
         Width           =   1245
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4530
      Top             =   270
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmPlanoinspecao_Tipodimensao.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   6
      Top             =   5340
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
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   5
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
End
Attribute VB_Name = "frmPlanoinspecao_Tipodimensao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Plano_Tipo_Dimensao As Boolean 'OK

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

If Qualidade_Plano = True Then
    NomeTabela = "Tipodimensao"
    CampoFiltro = "idtipo"
    CampoFiltro1 = "tipo"
    TextoFiltro = ""
Else
    NomeTabela = "Embalagem_Marca_Especie"
    CampoFiltro = "ID"
    CampoFiltro1 = "Descricao"
    If Aplic = 1 Then TextoFiltro = "where tipo = 'E'" Else TextoFiltro = "where tipo = 'M'"
End If
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CampoFiltro & " as ID, " & CampoFiltro1 & " as Descricao from " & NomeTabela & " " & TextoFiltro & " order by " & CampoFiltro1, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
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

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If Qualidade_Plano = True Then
                    If USMsgBox("Deseja realmente excluir este(s) tipos(s) da dimensão?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                Else
                    If Aplic = 1 Then TextoMsg = "espécie" Else TextoMsg = "marca"
                    If USMsgBox("Deseja realmente excluir esta(s) " & TextoMsg & "(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                End If
            End If
1:
            Permitido = True
            
            If Qualidade_Plano = True Then
                Conexao.Execute "DELETE from Tipodimensao where idtipo = " & .ListItems(InitFor)
                Modulo = "Qualidade/Plano de inspeção"
                Evento = "Excluir tipo da dimensão"
                Documento = "Tipo da dimensão: " & .ListItems(InitFor).ListSubItems(1)
            Else
                Modulo = Formulario
                Conexao.Execute "DELETE from Embalagem_Marca_Especie where ID = " & .ListItems(InitFor)
                If Aplic = 1 Then
                    Evento = "Excluir espécie"
                    Documento = "Espécie: " & .ListItems(InitFor).ListSubItems(1)
                Else
                    Evento = "Excluir marca"
                    Documento = "Marca: " & .ListItems(InitFor).ListSubItems(1)
                End If
            End If
            '==================================
            ID_documento = .ListItems(InitFor)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With

If Qualidade_Plano = True Then
    TextoFiltro = "o(s) tipo(s) da dimensão"
    TextoFiltro1 = "Tipo(s) da dimensão excluído(s)"
Else
    If Aplic = 1 Then
        TextoFiltro = "a(s) espécie(s)"
        TextoFiltro1 = "Espécie(s) excluída(s)"
    Else
        TextoFiltro = "a(s) marca(s)"
        TextoFiltro1 = "Marca(s) excluída(s)"
    End If
End If

If Permitido = False Then
    USMsgBox ("Informe " & TextoFiltro & "  antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox (TextoFiltro1 & " com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    Frame1.Enabled = False
    ProcCarregaLista
    Novo_Plano_Tipo_Dimensao = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Plano_Tipo_Dimensao = True
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
    If Qualidade_Plano = True Then
        NomeCampo = "o tipo da dimensão"
        Acao = "salvar"
        ProcVerificaAcao
        txtTexto.SetFocus
        Exit Sub
    End If
End If
If Qualidade_Plano = True Then
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Tipodimensao where idtipo = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then TBGravar.AddNew
    TBGravar!Tipo = txtTexto.Text
    TBGravar.Update
    txtId = TBGravar!IDTipo
    TBGravar.Close
    ProcCarregaLista
    If Novo_Plano_Tipo_Dimensao = True Then
        USMsgBox ("Novo tipo da dimensão cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Nova dimensão"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar tipo da dimensão"
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
    '==================================
    Modulo = "Qualidade/Plano de inspeção"
    ID_documento = txtId
    Documento = "Tipo da dimensão: " & txtTexto
    Documento1 = ""
    ProcGravaEvento
    '==================================
Else
    If Aplic = 1 Then
        TipoTexto = "E"
        TextoMsg = "epécie"
        Documento = "Espécie: " & txtTexto
    Else
        TipoTexto = "M"
        TextoMsg = "marca"
        Documento = "Marca: " & txtTexto
    End If
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Embalagem_Marca_Especie where id = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
    Else
        If txtTexto <> TBGravar!Descricao Then
            If Aplic = 1 Then TextoFiltro = "txt_Marca = '" & txtTexto & "' where txt_Marca = '" & TBGravar!Descricao & "'" Else TextoFiltro = "txt_Especie = '" & txtTexto & "' where txt_Especie = '" & TBGravar!Descricao & "'"
            Conexao.Execute "UPDATE tbl_Dados_Transp Set " & TextoFiltro
        End If
    End If
    TBGravar!Descricao = txtTexto.Text
    TBGravar!Tipo = TipoTexto
    TBGravar.Update
    txtId = TBGravar!ID
    TBGravar.Close
    ProcCarregaLista
    If Novo_Plano_Tipo_Dimensao = True Then
        USMsgBox ("Nova " & TextoMsg & " cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Nova " & TextoMsg
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar " & TextoMsg
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
    '==================================
    Modulo = Formulario
    ID_documento = txtId
    Documento1 = ""
    ProcGravaEvento
    '==================================
    With frmFaturamento_Prod_Serv
        If Aplic = 1 Then ProcCarregaComboTranspMarcaEspecie .Cmb_transp_especie, "E" Else ProcCarregaComboTranspMarcaEspecie .Cmb_transp_marca, "M"
    End With
End If
Novo_Plano_Tipo_Dimensao = False

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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7455, 7, True

If Aplic = 1 Then Familiatext = "Espécie" Else Familiatext = "Marca"
If Qualidade_Plano = False Then
    If Formulario = "Faturamento/Nota fiscal/Própria" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Própria - " & Familiatext
    ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
            Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - " & Familiatext
        ElseIf Formulario = "Estoque/Ordem de faturamento" Then
                Caption = "Estoque - Ordem de faturamento - " & Familiatext
            Else
                Caption = "Estoque - Nota fiscal - " & Familiatext
    End If
    Label1.Caption = Familiatext
    txtTexto.ToolTipText = Familiatext & "."
    Lista.ColumnHeaders.Item(2).Text = Familiatext
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

If Novo_Plano_Tipo_Dimensao = True Then
    If Qualidade_Plano = True Then
        TextoFiltro = "O tipo da dimensão ainda não foi salvo"
    Else
        If Aplic = 1 Then TextoFiltro = "A espécie ainda não foi salva" Else TextoFiltro = "A marca ainda não foi salva"
    End If
    If USMsgBox(TextoFiltro & ", deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Plano_Tipo_Dimensao = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
    Novo_Plano_Tipo_Dimensao = False
    Unload Me
    If Qualidade_Plano = True Then frmPlanoinspecao.ProcCarrega_Tipo
End If
Novo_Plano_Tipo_Dimensao = False
Unload Me

If Qualidade_Plano = False Then
    With frmFaturamento_Prod_Serv
        If Aplic = 1 Then
            ProcCarregaComboTranspMarcaEspecie .Cmb_transp_especie, "E"
            .Cmb_transp_especie.Refresh
        Else
            ProcCarregaComboTranspMarcaEspecie .Cmb_transp_marca, "M"
            .Cmb_transp_marca.Refresh
        End If
    End With
End If

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
                If Qualidade_Plano = False Then
                    If Aplic = 1 Then
                        ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Transp", "txt_Especie = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                    Else
                        ProcVerificaRegistroUtilizadoSemMsg "tbl_Dados_Transp", "txt_Marca = '" & .ListItems(InitFor).ListSubItems(1) & "'"
                        If Permitido = False Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                    End If
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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
If Qualidade_Plano = True Then
    TBLISTA.Open "Select * from Tipodimensao where IDtipo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        With frmPlanoinspecao
            .ProcCarrega_Tipo
            .cmbtipomed = TBLISTA!Tipo
        End With
    End If
Else
    TBLISTA.Open "Select * from Embalagem_Marca_Especie where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
    
    If Formulario <> "Estoque/Ordem de faturamento" Then
        With frmFaturamento_Prod_Serv
            If Aplic = 1 Then .Cmb_transp_especie = TBLISTA!Descricao Else .Cmb_transp_marca = TBLISTA!Descricao
        End With
    Else
        With frmEstoque_Ordem_Faturamento
            If Aplic = 1 Then .Cmb_transp_especie = TBLISTA!Descricao Else .Cmb_transp_marca = TBLISTA!Descricao
        End With
    End If
    
    End If
End If
TBLISTA.Close
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
            If Qualidade_Plano = False Then
                If Aplic = 1 Then
                    Mensagem = "Não é permitido excluir esta espécie, pois a mesma está sendo utilizada no módulo"
                    ProcVerificaRegistroUtilizado "tbl_Dados_Transp", "txt_Especie = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Faturamento/Nota fiscal"
                    If Permitido = False Then .ListItems.Item(InitFor).Checked = False
                Else
                    Mensagem = "Não é permitido excluir esta marca, pois a mesma está sendo utilizada no módulo"
                    ProcVerificaRegistroUtilizado "tbl_Dados_Transp", "txt_Marca = '" & .ListItems(InitFor).ListSubItems(1) & "'", "Faturamento/Nota fiscal"
                    If Permitido = False Then .ListItems.Item(InitFor).Checked = False
                End If
            End If
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
ProcLimpaCampos
If Qualidade_Plano = True Then
    NomeTabela = "Tipodimensao"
    NomeCampo = "Idtipo"
Else
    NomeTabela = "Embalagem_marca_especie"
    NomeCampo = "ID"
End If
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select * from " & NomeTabela & " where " & NomeCampo & " = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    ProcCarregaDados
    CodigoLista = Lista.SelectedItem.index
End If
TBplano.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

If Qualidade_Plano = True Then
    txtId = TBplano!IDTipo
    txtTexto = IIf(IsNull(TBplano!Tipo), "", TBplano!Tipo)
Else
    txtId.Text = TBplano!ID
    txtTexto.Text = TBplano!Descricao
End If
Frame1.Enabled = True
Novo_Plano_Tipo_Dimensao = False
    
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
