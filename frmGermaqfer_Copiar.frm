VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGermaqfer_Copiar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerenciamento de ordem - Menu impressão"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   3165
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePosto 
      BackColor       =   &H00E0E0E0&
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
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   990
      Width           =   8655
      Begin VB.ComboBox cmbPosto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Número da ordem."
         Top             =   390
         Width           =   2325
      End
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   5955
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   1095
         TabIndex        =   5
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   5152
         TabIndex        =   4
         Top             =   180
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2145
      Left            =   60
      TabIndex        =   6
      Top             =   975
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   3784
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
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
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Mês"
         Object.Width           =   5054
      EndProperty
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5010
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmGermaqfer_Copiar.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
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
      ButtonCaption1  =   "Copiar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Copiar (F3)"
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
      ButtonWidth1    =   44
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
      ButtonLeft2     =   48
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   52
      ButtonTop3      =   2
      ButtonWidth3    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   95
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   127
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
End
Attribute VB_Name = "frmGermaqfer_Copiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbPosto_Click()
On Error GoTo tratar_erro

txtdescricao = ""
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select Descricao from Cadmaquinas where maquina = '" & cmbPosto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then txtdescricao = IIf(IsNull(TBMaquinas!Descricao), "", TBMaquinas!Descricao)
TBMaquinas.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcCopiar
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Sit_REG = 1 Then
    ProcCarregaToolBar1 Me, 8655, 5, True
    Height = 2295
    Lista.Visible = False
    
    ProcCarregaComboPostoTrab cmbPosto, "maquina <> '" & frmGermaqfer.txtmaquina & "' and Maquina is not null", False, False
Else
    ProcCarregaToolBar1 Me, 3555, 5, True
    Width = 3765
    FramePosto.Visible = False
    
    With Lista.ListItems
        .Clear
        Contador = 1
        Do While Contador <= 7
            .Add , , ""
            Select Case Contador
                Case 1: Texto = "Domingo"
                Case 2: Texto = "Segunda"
                Case 3: Texto = "Terça"
                Case 4: Texto = "Quarta"
                Case 5: Texto = "Quinta"
                Case 6: Texto = "Sexta"
                Case 7: Texto = "Sábado"
            End Select
            .Item(.Count).SubItems(1) = Texto
            Contador = Contador + 1
        Loop
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
                If frmGermaqfer.cmbdia = .ListItems(InitFor).SubItems(1) Then GoTo Proximo

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
            If frmGermaqfer.cmbdia = .ListItems(InitFor).SubItems(1) Then
                USMsgBox ("Não é permitido copiar para o mesmo dia da semana do turno selecionado."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcCopiar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarTurno_Posto()
On Error GoTo tratar_erro

If cmbPosto = "" Then
    USMsgBox ("Informe o posto de trabalho antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If USMsgBox("Deseja realmente copiar os dados do(s) turno(s) deste posto de trabalho?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select IDMaquina from Cadmaquinas where maquina = '" & cmbPosto & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        IDMaquina = TBFIltro!IDMaquina
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from CadmaqTurnos where maquina = '" & frmGermaqfer.txtmaquina & "' order by diasemana, turno", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            Do While TBMaquinas.EOF = False
                Set TBGravar = CreateObject("adodb.recordset")
                TBGravar.Open "Select * from CadmaqTurnos where maquina = '" & cmbPosto & "' and diasemana = '" & TBMaquinas!Diasemana & "' and turno = " & TBMaquinas!Turno, Conexao, adOpenKeyset, adLockOptimistic
                If TBGravar.EOF = True Then TBGravar.AddNew
                TBGravar!maquina = cmbPosto
                TBGravar!Bloqueado = False
                TBGravar!Data = Date
                TBGravar!Responsavel = pubUsuario
                TBGravar!Diasemana = TBMaquinas!Diasemana
                TBGravar!Turno = TBMaquinas!Turno
                TBGravar!Inicioturno = TBMaquinas!Inicioturno
                TBGravar!Margem_inicio_ap = TBMaquinas!Margem_inicio_ap
                TBGravar!finalturno = TBMaquinas!finalturno
                TBGravar!Total = TBMaquinas!Total
                TBGravar!intervalo = TBMaquinas!intervalo
                TBGravar!Inicio_intervalo = TBMaquinas!Inicio_intervalo
                TBGravar!Final_intervalo = TBMaquinas!Final_intervalo
                TBGravar!TotalTurno = TBMaquinas!TotalTurno
                TBGravar!TotalDia = TBMaquinas!TotalDia
                TBGravar.Update
                TBGravar.Close
                '==================================
                Modulo = "PCP/Postos de trabalho"
                Evento = "Novo turno"
                ID_documento = IDMaquina
                Documento = "Código do posto de trabalho: " & cmbPosto
                Documento1 = "Dia da semana: " & TBMaquinas!Diasemana & " - Truno: " & TBMaquinas!Turno
                ProcGravaEvento
                '==================================
                TBMaquinas.MoveNext
            Loop
        Else
            TBMaquinas.Close
            Exit Sub
        End If
        TBMaquinas.Close
        USMsgBox ("Dados do(s) turno(s) copiado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
    TBFIltro.Close
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiarTurno()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente copiar os dados do turno para o(s) dia(s) da semana selecionado(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Permitido = False
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from CadmaqTurnos where maquina = '" & frmGermaqfer.txtmaquina & "' and Diasemana = '" & frmGermaqfer.cmbdia & "' order by diasemana, turno", Conexao, adOpenKeyset, adLockOptimistic
                Do While TBMaquinas.EOF = False
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from CadmaqTurnos where maquina = '" & frmGermaqfer.txtmaquina & "' and diasemana = '" & .ListItems(InitFor).SubItems(1) & "' and turno = " & TBMaquinas!Turno, Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then TBGravar.AddNew
                    TBGravar!maquina = TBMaquinas!maquina
                    TBGravar!Bloqueado = False
                    TBGravar!Data = Date
                    TBGravar!Responsavel = pubUsuario
                    TBGravar!Diasemana = .ListItems(InitFor).SubItems(1)
                    TBGravar!Turno = TBMaquinas!Turno
                    TBGravar!Inicioturno = TBMaquinas!Inicioturno
                    TBGravar!Margem_inicio_ap = TBMaquinas!Margem_inicio_ap
                    TBGravar!finalturno = TBMaquinas!finalturno
                    TBGravar!Total = TBMaquinas!Total
                    TBGravar!intervalo = TBMaquinas!intervalo
                    TBGravar!Inicio_intervalo = TBMaquinas!Inicio_intervalo
                    TBGravar!Final_intervalo = TBMaquinas!Final_intervalo
                    TBGravar!TotalTurno = TBMaquinas!TotalTurno
                    TBGravar!TotalDia = TBMaquinas!TotalDia
                    TBGravar.Update
                    TBGravar.Close
                    '==================================
                    Modulo = "PCP/Postos de trabalho"
                    Evento = "Novo turno"
                    ID_documento = frmGermaqfer.txtIDmaquina
                    Documento = "Código do posto de trabalho: " & frmGermaqfer.txtmaquina
                    Documento1 = "Dia da semana: " & .ListItems(InitFor).SubItems(1) & " - Truno: " & TBMaquinas!Turno
                    ProcGravaEvento
                    '==================================
                    TBMaquinas.MoveNext
                Loop
                TBMaquinas.Close
                Permitido = True
            End If
        Next InitFor
    End With
    If Permitido = False Then
        USMsgBox ("Informe o(s) dia(s) da semana antes de copiar o(s) turno(s)."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Turno(s) copiado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
        frmGermaqfer.ProcCarregaTurnos
        Unload Me
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

If Sit_REG = 1 Then ProcCopiarTurno_Posto Else ProcCopiarTurno

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
