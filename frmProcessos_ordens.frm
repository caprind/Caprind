VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessos_ordens 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Processos - Validar tempos da fase"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11910
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
   Icon            =   "frmProcessos_ordens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   8760
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmProcessos_ordens.frx":1042
      Count           =   1
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5190
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9155
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483628
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Fase"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Gupo/op."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Posto de trab."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   1914
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Ct. hr. prep."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "D"
         Text            =   "Prep. utiliz. pç"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Ct. hr. exe."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Exec. utiliz. pç"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Qtde. OK"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
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
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "4"
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
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   6210
      Width           =   11775
      _ExtentX        =   20770
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
Attribute VB_Name = "frmProcessos_ordens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
If USMsgBox("Deseja realmente validar os tempos da fase?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    With frmProcessos
        If FunVerificaRegistroValidado("Processos", "IDProcesso = " & .txtidprocesso, "processo", "os tempos da fase", "validar", True, True) = False Then Exit Sub
        
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select * from fases where idfase = " & .ListaFases.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = False Then
            TBCFOP!Preparacao = Lista.SelectedItem.ListSubItems.Item(7).Text
            ProcFormataHora (Lista.SelectedItem.ListSubItems.Item(9).Text)
            TBCFOP!Execucao = DataResultado
            TBCFOP!TempoPreparacao = Lista.SelectedItem.ListSubItems.Item(7).Text
            TBCFOP!TempoExecucao = Lista.SelectedItem.ListSubItems.Item(9).Text
            TBCFOP!cronometrado = True
            TBCFOP.Update
            USMsgBox ("Tempos da fase validados com sucesso."), vbInformation, "CAPRIND v5.0"
        
            '==================================
            Evento = "Validar tempos da fase"
            Modulo = "Engenharia/Processos/Atualizar fase do processo"
            ID_documento = .ListaFases.SelectedItem
            Documento = "Processo: " & .txtidprocesso & " - Rev.: " & .txtrevproc & " - Cód. interno: " & .txtdesenho & " - Rev.: " & .txtrevdesenho
            Documento1 = "Fase: " & .txtFase & " - Grupo/op.: " & .txtgrupo_op & " - Posto: " & .cmbMaquina
            ProcGravaEvento
            '==================================
            .ProcAtualizaFases
            .ProcPuxaDadosFase
            Unload Me
        Else
            USMsgBox ("Não existe cadastro de fase para esta ordem de serviço."), vbInformation, "CAPRIND v5.0"
        End If
        TBCFOP.Close
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
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

ProcCarregaToolBar1 Me, 11775, 5, True
ProcCarregaLista
    
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

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Controle = Year(Date) - 1
If Format(Date, "dd/mm") = "29/02" Then
    If FunBissexto(Controle) = False Then Dataini = "28/02/" & Controle Else Dataini = Format(Date, "dd/mm") & "/" & Controle
Else
    Dataini = Format(Date, "dd/mm") & "/" & Controle
End If
Set TBOrdem = CreateObject("adodb.recordset")
NomeCamposFiltro = "OS.IDProducao, OS.Ordem, OS.Fase, OS.Grupo_op, OS.maquina, OS.TPUTIL, OS.TPUSEG, OS.TEUTIL, OS.TEUSEG, OS.QTOK, M.Descricao, OS.Valor_hs_prep, OS.Valor_hs_exec"
TBOrdem.Open "Select " & NomeCamposFiltro & " from Ordemservico OS INNER JOIN CadMaquinas M on OS.Maquina = M.Maquina where OS.IDFase = " & frmProcessos.ListaFases.SelectedItem & " and Year(OS.Prazofinal) >= '" & Year(Dataini) & "' order by OS.fase, OS.retrabalho, OS.IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    TBOrdem.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBOrdem.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBOrdem.MoveFirst
    Do While TBOrdem.EOF = False
        With Lista.ListItems
            .Add , , IIf(IsNull(TBOrdem!Ordem), "", TBOrdem!Ordem)
            .Item(.Count).SubItems(1) = IIf(IsNull(TBOrdem!IDProducao), "", TBOrdem!IDProducao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBOrdem!Fase), "", TBOrdem!Fase)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBOrdem!Grupo_op), "", TBOrdem!Grupo_op)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBOrdem!maquina), "", TBOrdem!maquina)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBOrdem!Descricao), "", TBOrdem!Descricao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBOrdem!Valor_hs_prep), "", Format(TBOrdem!Valor_hs_prep, "###,##0.00"))
                       
            'Tempo de preparação real p/ peça
            TPUSEG = IIf(IsNull(TBOrdem!TPUSEG), 0, TBOrdem!TPUSEG)
            s = TPUSEG
            .Item(.Count).SubItems(7) = FormataTempo(s)
            
            .Item(.Count).SubItems(8) = IIf(IsNull(TBOrdem!Valor_hs_exec), "", Format(TBOrdem!Valor_hs_exec, "###,##0.00"))
            
            'Tempo de execucao real p/ peça
            .Item(.Count).SubItems(9) = IIf(IsNull(TBOrdem!TEUTIL), "00:00:00", Format(TBOrdem!TEUTIL, "hh:mm:ss"))
            
            .Item(.Count).SubItems(10) = IIf(IsNull(TBOrdem!QTOK), "", Format(TBOrdem!QTOK, "###,##0.00"))
        End With
        TBOrdem.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count <> 0 Then
    If Lista.SelectedItem = "" Then Exit Sub
    Formulario = "PCP/Gerenciamento de ordem"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    Unload Me
    With frmprod
        .SSTab1.Tab = 1
        Ordem = Lista.SelectedItem
        .ProcCarregaOrdem
    End With
End If

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
