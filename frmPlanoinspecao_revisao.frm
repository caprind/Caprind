VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanoinspecao_revisao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Plano de inspe��o - Hist�rico de revis�es"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
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
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.TextBox txtID 
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
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "0"
      ToolTipText     =   "ID"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListaHistorico 
      Height          =   6855
      Left            =   60
      TabIndex        =   4
      Top             =   2895
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12091
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
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
         Text            =   "Respons�vel"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Hist�rico"
         Object.Width           =   17136
      EndProperty
   End
   Begin VB.Frame Frame3 
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
      Height          =   1905
      Left            =   55
      TabIndex        =   5
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtAnalisado 
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
         Left            =   2790
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Respons�vel pela revis�o."
         Top             =   390
         Width           =   12195
      End
      Begin VB.TextBox txtRevisao 
         Alignment       =   2  'Center
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
         Left            =   180
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "N�mero da revis�o."
         Top             =   390
         Width           =   1175
      End
      Begin VB.TextBox txtHistorico 
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
         Height          =   765
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "Hist�rico da revis�o."
         Top             =   990
         Width           =   14805
      End
      Begin MSComCtl2.DTPicker txtDtRevisao 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         ToolTipText     =   "Data da revis�o."
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   143065089
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Respons�vel pela revis�o"
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
         Left            =   7972
         TabIndex        =   9
         Top             =   180
         Width           =   1830
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Revis�o"
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
         Height          =   195
         Left            =   430
         TabIndex        =   8
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Hist�rico"
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
         Left            =   7275
         TabIndex        =   7
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. revis�o"
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
         Height          =   255
         Left            =   1635
         TabIndex        =   6
         Top             =   180
         Width           =   855
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   11
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   9
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
      ButtonCaption4  =   "Anterior"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Registro anterior."
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
      ButtonLeft4     =   118
      ButtonTop4      =   2
      ButtonWidth4    =   47
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Pr�ximo"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Pr�ximo registro."
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
      ButtonLeft5     =   167
      ButtonTop5      =   2
      ButtonWidth5    =   46
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonAlignment6=   2
      ButtonType6     =   1
      ButtonStyle6    =   -1
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   -1
      ButtonLeft6     =   215
      ButtonTop6      =   4
      ButtonWidth6    =   2
      ButtonHeight6   =   54
      ButtonCaption7  =   "Ajuda"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Ajuda (F1)"
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
      ButtonLeft7     =   219
      ButtonTop7      =   2
      ButtonWidth7    =   36
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Sair"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Sair (Esc)"
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   257
      ButtonTop8      =   2
      ButtonWidth8    =   26
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   5
      ButtonLeft9     =   285
      ButtonTop9      =   2
      ButtonWidth9    =   24
      ButtonHeight9   =   24
      ButtonUseMaskColor9=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11640
         Top             =   195
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmPlanoinspecao_revisao.frx":0000
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   12
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
End
Attribute VB_Name = "frmPlanoinspecao_revisao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Plano_Rev_Historico As Boolean 'OK

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Plano_revisao where idplano = " & frmPlanoinspecao.Lista.SelectedItem & " order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId = TBLISTA!ID
        Set TBHistProc = CreateObject("adodb.recordset")
        TBHistProc.Open "Select * from Plano_revisao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBHistProc.EOF = False Then
            ProcLimpaCampos
            ProcCarregaDados
        End If
    Else
        USMsgBox ("Fim dos cadastros do hist�rico de revis�o."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Plano_revisao where idplano = " & frmPlanoinspecao.Lista.SelectedItem & " order by id", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("id = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId = TBLISTA!ID
        Set TBHistProc = CreateObject("adodb.recordset")
        TBHistProc.Open "Select * from Plano_revisao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBHistProc.EOF = False Then
            ProcLimpaCampos
            ProcCarregaDados
        End If
    Else
        USMsgBox ("Fim dos cadastros do hist�rico de revis�o."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista()
On Error GoTo tratar_erro

ListaHistorico.ListItems.Clear
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Plano_revisao where idplano = " & frmPlanoinspecao.txtPI & " order by revisao", Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    TBHistProc.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBHistProc.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBHistProc.MoveFirst
    Do While TBHistProc.EOF = False
        frmPlanoinspecao.txtRev = IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao)
        With ListaHistorico.ListItems
            .Add , , TBHistProc!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBHistProc!Revisao), "", (TBHistProc!Revisao))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBHistProc!Data), "", Format(TBHistProc!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBHistProc!por), "", (TBHistProc!por))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBHistProc!Historico), "", (TBHistProc!Historico))
        End With
        TBHistProc.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBHistProc.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 9, True
Formulario = "Qualidade/Plano de inspe��o/Hist�rico de revis�es"
txtDtRevisao.Value = Date
ProcLimpaVariaveisPrincipais
ProcAtualizalista

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Plano de inspe��o/Hist�rico de revis�es"
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With frmPlanoinspecao
    For InitFor = 1 To ListaHistorico.ListItems.Count
        If ListaHistorico.ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) revis�o(�es)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Plano_revisao where id = " & ListaHistorico.ListItems(InitFor)
            '==================================
            Modulo = Formulario
            Evento = "Excluir"
            ID_documento = ListaHistorico.ListItems(InitFor)
            Documento = "Plano: " & .txtPI & " - Rev.: " & .txtRev & " - C�d. interno: " & .txtdesenho & " - Rev.: " & .txtRev_item
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
    
    If Permitido = False Then
        USMsgBox ("Informe a(s) revis�o(�es) antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Revis�o(�es) exclu�da(s) com sucesso."), vbInformation, "CAPRIND v5.0"
        Set TBplano = CreateObject("adodb.recordset")
        TBplano.Open "Select * from plano where IDplano = " & .txtPI, Conexao, adOpenKeyset, adLockOptimistic
        If TBplano.EOF = False Then
            Set TBHistProc = CreateObject("adodb.recordset")
            TBHistProc.Open "Select revisao from Plano_revisao where idplano = " & .txtPI & " order by Revisao", Conexao, adOpenKeyset, adLockOptimistic
            If TBHistProc.EOF = False Then
                TBHistProc.MoveLast
                TBplano!Rev = TBHistProc!Revisao
            Else
                TBplano!Rev = 0
            End If
            TBHistProc.Close
            TBplano.Update
            .ProcCarregaDados
            .Lista.ListItems.Clear
            .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
        End If
        TBplano.Close
        
        ProcLimpaCampos
        ProcAtualizalista
        Frame3.Enabled = False
        Novo_Plano_Rev_Historico = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

With frmPlanoinspecao
    If .txtIDFase <> "" And .txtIDFase <> "0" Then
        If .FunVerifiProcRevisado(.txtIDFase, "criar nova revis�o deste plano", True) = True Then Exit Sub
    End If
End With
ProcLimpaCampos
Novo_Plano_Rev_Historico = True
Frame3.Enabled = True
txtrevisao.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtrevisao.Text = ""
txtDtRevisao.Value = Date
txtAnalisado.Text = ""
txtHistorico.Text = ""
CodigoLista = 0
Caption = "Qualidade/Plano de inspe��o/Hist�rico de revis�es"

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Plano_Rev_Historico = True Then
    If USMsgBox("O hist�rico ainda n�o foi salvo, deseja salvar antes de fechar o m�dulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Plano_Rev_Historico = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Plano_Rev_Historico = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtrevisao = "" Then
    NomeCampo = "a revis�o"
    ProcVerificaAcao
    txtrevisao.SetFocus
    Exit Sub
End If
If txtAnalisado = "" Then
    NomeCampo = "o respons�vel"
    ProcVerificaAcao
    txtAnalisado.SetFocus
    Exit Sub
End If
If txtHistorico = "" Then
    NomeCampo = "o hist�rico"
    ProcVerificaAcao
    txtHistorico.SetFocus
    Exit Sub
End If
With frmPlanoinspecao
    Set TBHistProc = CreateObject("adodb.recordset")
    TBHistProc.Open "Select * from Plano_revisao where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBHistProc.EOF = True Then
        TBHistProc.AddNew
    Else
        If FunVerifValidacaoRegistro("alterar", .txtDtValidacao, "plano", "a revis�o do plano", True) = False Then Exit Sub
        With frmPlanoinspecao
            If .txtIDFase <> "" And .txtIDFase <> "0" Then
                If .FunVerifiProcRevisado(.txtIDFase, "alterar a revis�o deste plano", True) = True Then Exit Sub
            End If
        End With
    End If
    TBHistProc!IDPlano = .txtPI
    TBHistProc!Data = txtDtRevisao
    TBHistProc!Revisao = txtrevisao
    TBHistProc!por = txtAnalisado
    TBHistProc!Historico = txtHistorico
    TBHistProc.Update
    txtId = TBHistProc!ID
    TBHistProc.Close
    ProcAtualizalista
    If Novo_Plano_Rev_Historico = True Then
        USMsgBox ("Nova revis�o cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
    Else
        USMsgBox ("Altera��o efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar"
        If CodigoLista <> 0 And ListaHistorico.ListItems.Count <> 0 Then
            ListaHistorico.SelectedItem = ListaHistorico.ListItems(CodigoLista)
            ListaHistorico.SetFocus
        End If
    End If
    '==================================
    Modulo = Formulario
    ID_documento = txtId
    Documento = "Plano: " & .txtPI & " - Rev.: " & .txtRev & " - C�d. interno: " & .txtdesenho & " - Rev.: " & .txtRev_item
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Novo_Plano_Rev_Historico = False
    
    Set TBplano = CreateObject("adodb.recordset")
    TBplano.Open "Select * from plano where IDplano = " & .txtPI, Conexao, adOpenKeyset, adLockOptimistic
    If TBplano.EOF = False Then
        TBplano!Rev = txtrevisao
        TBplano.Update
        .ProcCarregaDados
        .Lista.ListItems.Clear
        .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    End If
    TBplano.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaHistorico_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaHistorico
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Plano", "IDPlano = " & frmPlanoinspecao.txtPI, True) = False Then GoTo Proximo
                If frmPlanoinspecao.txtIDFase <> "" And frmPlanoinspecao.txtIDFase <> "0" Then
                    If frmPlanoinspecao.FunVerifiProcRevisado(frmPlanoinspecao.txtIDFase, "excluir a revis�o deste plano", False) = True Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaHistorico, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaHistorico_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaHistorico
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Plano", "IDPlano = " & frmPlanoinspecao.txtPI, "plano", "revis�o do plano", "excluir esta", True, True) = False Then
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If frmPlanoinspecao.txtIDFase <> "" And frmPlanoinspecao.txtIDFase <> "0" Then
                If frmPlanoinspecao.FunVerifiProcRevisado(frmPlanoinspecao.txtIDFase, "excluir a revis�o deste plano", True) = True Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaHistorico_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If ListaHistorico.ListItems.Count = 0 Then Exit Sub
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from Plano_revisao where id = " & ListaHistorico.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = ListaHistorico.SelectedItem.index
End If
TBHistProc.Close
Frame3.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId = TBHistProc!ID
Caption = "Qualidade - Plano de inspe��o - Hist�rico de revis�es - (Revis�o : " & IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao) & ")"
txtrevisao.Text = IIf(IsNull(TBHistProc!Revisao), "", TBHistProc!Revisao)
txtAnalisado.Text = IIf(IsNull(TBHistProc!por), "", (TBHistProc!por))
txtHistorico = IIf(IsNull(TBHistProc!Historico), "", (TBHistProc!Historico))
If IsNull(TBHistProc!Data) = False Then txtDtRevisao.Value = TBHistProc!Data
Novo_Plano_Rev_Historico = False

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: ProcAnterior
    Case 5: ProcProximo
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
