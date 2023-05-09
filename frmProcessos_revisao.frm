VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessos_revisao 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Engenharia - Processos - Histórico de revisões"
   ClientHeight    =   10035
   ClientLeft      =   120
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
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
      MouseIcon       =   "frmProcessos_revisao.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Text            =   "0"
      ToolTipText     =   "ID"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
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
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Analisado e aprovado."
         Top             =   390
         Width           =   12210
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
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número de revisão."
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
         ToolTipText     =   "Histórico da revisão."
         Top             =   990
         Width           =   14775
      End
      Begin MSComCtl2.DTPicker txtDtRevisao 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         ToolTipText     =   "Data da revisão."
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
         Format          =   198049793
         CurrentDate     =   39057
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Analisado/aprovado*"
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
         Left            =   8100
         TabIndex        =   9
         Top             =   180
         Width           =   1530
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Revisão"
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
         Caption         =   "Histórico*"
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
         Left            =   7215
         TabIndex        =   7
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dt. revisão"
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
   Begin MSComctlLib.ListView ListaHistorico 
      Height          =   6830
      Left            =   60
      TabIndex        =   4
      Top             =   2910
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12039
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Text            =   "ID"
         Object.Width           =   0
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
         Text            =   "Responsável"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Histórico"
         Object.Width           =   17136
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   8
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
      ButtonCaption3  =   "Anterior"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Registro anterior."
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
      ButtonWidth3    =   47
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Próximo"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Próximo registro."
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
      ButtonLeft4     =   126
      ButtonTop4      =   2
      ButtonWidth4    =   46
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   174
      ButtonTop5      =   4
      ButtonWidth5    =   2
      ButtonHeight5   =   54
      ButtonCaption6  =   "Ajuda"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Ajuda (F1)"
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
      ButtonLeft6     =   178
      ButtonTop6      =   2
      ButtonWidth6    =   36
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
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
      ButtonLeft7     =   216
      ButtonTop7      =   2
      ButtonWidth7    =   26
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
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
      ButtonState8    =   5
      ButtonLeft8     =   244
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      ButtonUseMaskColor8=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11640
         Top             =   195
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProcessos_revisao.frx":030A
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
Attribute VB_Name = "frmProcessos_revisao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Processo_Historico As Boolean 'OK

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from histproc where idprocesso = " & frmProcessos.txtidprocesso & " order by Ordem", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDHistorico = " & txtId)
    TBLISTA.MovePrevious
    If TBLISTA.BOF = False Then
        txtId = TBLISTA!IDHistorico
        Set TBHistProc = CreateObject("adodb.recordset")
        TBHistProc.Open "Select * from histproc where IDHistorico = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBHistProc.EOF = False Then
            ProcLimpaCampos
            ProcCarregaDados
        End If
    Else
        USMsgBox ("Fim dos cadastros do histórico de revisão."), vbInformation, "CAPRIND v5.0"
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtId = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from histproc where idprocesso = " & frmProcessos.txtidprocesso & " order by Ordem", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.BOF = False Then
    TBLISTA.Find ("IDHistorico = " & txtId)
    TBLISTA.MoveNext
    If TBLISTA.EOF = False Then
        txtId = TBLISTA!IDHistorico
        Set TBHistProc = CreateObject("adodb.recordset")
        TBHistProc.Open "Select * from histproc where IDHistorico = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBHistProc.EOF = False Then
            ProcLimpaCampos
            ProcCarregaDados
        End If
    Else
        USMsgBox ("Fim dos cadastros do histórico de revisão."), vbInformation, "CAPRIND v5.0"
    End If
End If

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
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista()
On Error GoTo tratar_erro

ListaHistorico.ListItems.Clear
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from histproc where idprocesso = " & frmProcessos.txtidprocesso & " order by Ordem", Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    TBHistProc.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBHistProc.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBHistProc.MoveFirst
    Do While TBHistProc.EOF = False
        frmProcessos.txtrevproc.Text = IIf(IsNull(TBHistProc!Ordem), "", TBHistProc!Ordem)
        With ListaHistorico.ListItems
            .Add , , TBHistProc!IDHistorico
            .Item(.Count).SubItems(1) = IIf(IsNull(TBHistProc!Ordem), "", TBHistProc!Ordem)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBHistProc!Data), "", Format(TBHistProc!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBHistProc!por), "", TBHistProc!por)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBHistProc!Historico), "", TBHistProc!Historico)
        End With
        TBHistProc.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBHistProc.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
Formulario = "Engenharia/Processos"
Direitos
ProcLimpaVariaveisPrincipais

txtDtRevisao.Value = Date
ProcAtualizalista

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

With frmProcessos
    If FunVerificaRegistroValidado("Processos", "IDprocesso = " & .txtidprocesso, "mesmo", "deste processo", "criar revisão", True, False) = False Then Exit Sub
    ProcLimpaCampos
    Novo_Processo_Historico = True
    Frame3.Enabled = True
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select Revisao from Processos where IDprocesso = " & .txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        txtrevisao = TBProcessos!Revisao + 1
    End If
    TBProcessos.Close
    txtDtRevisao.SetFocus
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
Caption = "Engenharia - Processos - Gerenciamento - Histórico de revisões"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Processo_Historico = True Then
    If USMsgBox("O histórico ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Processo_Historico = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Processo_Historico = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtAnalisado = "" Then
    NomeCampo = "o responsável"
    ProcVerificaAcao
    txtAnalisado.SetFocus
    Exit Sub
End If
If txtHistorico = "" Then
    NomeCampo = "o histórico"
    ProcVerificaAcao
    txtHistorico.SetFocus
    Exit Sub
End If
If txtDtRevisao > Date Then
    USMsgBox ("A data da revisão não poder ser maior que a data atual."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

With frmProcessos
    Set TBHistProc = CreateObject("adodb.recordset")
    TBHistProc.Open "Select * from histproc where IDHistorico = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBHistProc.EOF = True Then TBHistProc.AddNew
    TBHistProc!IDPROCESSO = .txtidprocesso
    TBHistProc!Data = txtDtRevisao
    TBHistProc!Ordem = txtrevisao
    TBHistProc!por = txtAnalisado
    TBHistProc!Historico = txtHistorico
    TBHistProc.Update
    txtId = TBHistProc!IDHistorico
    TBHistProc.Close
    ProcAtualizalista
    If Novo_Processo_Historico = True Then
        USMsgBox ("Nova revisão cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Nova revisão"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar revisão"
        If CodigoLista <> 0 And ListaHistorico.ListItems.Count <> 0 Then
            ListaHistorico.SelectedItem = ListaHistorico.ListItems(CodigoLista)
            ListaHistorico.SetFocus
        End If
    End If
    '==================================
    Modulo = "Engenharia/Processos"
    ID_documento = txtId
    Documento = "Processo: " & .Txt_numero_processo & " - Rev.: " & .txtrevproc & " - Cód. interno: " & .txtdesenho & " - Rev.: " & .txtrevdesenho
    Documento1 = "Nº revisão: " & txtrevisao
    ProcGravaEvento
    '==================================
    
    If Novo_Processo_Historico = True Then
        'Cria novo processo com a nova revisão
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Processos where IDProcesso = " & .txtidprocesso, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBProcessos = CreateObject("adodb.recordset")
            TBProcessos.Open "Select * from Processos", Conexao, adOpenKeyset, adLockOptimistic
            TBProcessos.AddNew
            TBProcessos!Nprocesso = TBAbrir!Nprocesso
            TBProcessos!Revisao = txtrevisao
            TBProcessos!Tipo = TBAbrir!Tipo
            TBProcessos!Custo = TBAbrir!Custo
            TBProcessos!Custoprep = TBAbrir!Custoprep
            TBProcessos!Contador = TBAbrir!Contador
            TBProcessos!elaborado = pubUsuario
            TBProcessos!DtImplantacao = Date
            TBProcessos!Codproduto = TBAbrir!Codproduto
            TBProcessos!PcHora = TBAbrir!PcHora
            TBProcessos!TTotal = TBAbrir!TTotal
            TBProcessos!TTotalSEG = TBAbrir!TTotalSEG
            TBProcessos!Bloqueado = False
            TBProcessos!Ordenarprocesso = TBAbrir!Ordenarprocesso
    
            TBProcessos.Update
            IDPROCESSO = TBProcessos!IDPROCESSO
            TBProcessos.Close
        End If
        TBAbrir.Close
        .ProcCopiarFase "Select * from fases where idprocesso = " & .txtidprocesso, True, IDPROCESSO, .txtdesenho, .txtProduto, True
        
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select * from processos where IDProcesso = " & IDPROCESSO, Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessos.EOF = False Then
            .ProcPuxaDados
            .Sql_Processo_Localizar = "Select P.Desenho, P.Descricao, P.Unidade, P.Classe, P.Unidade, P.Classe, PR.* from projproduto P INNER JOIN Processos PR ON PR.codproduto = P.CodProduto where PR.IDProcesso = " & IDPROCESSO
            .ProcCarregaListaProcessos (1)
        End If
    End If
End With
Novo_Processo_Historico = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Engenharia/Processos"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaHistorico_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaHistorico, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaHistorico_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If ListaHistorico.ListItems.Count = 0 Then Exit Sub
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select * from histproc where IDHistorico = " & ListaHistorico.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = ListaHistorico.SelectedItem.index
End If
TBHistProc.Close
Frame3.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaDados()
On Error GoTo tratar_erro

txtId = TBHistProc!IDHistorico
Caption = "Engenharia - Processos - Gerenciamento - Histórico de revisões - (Revisão : " & IIf(IsNull(TBHistProc!Ordem), "", TBHistProc!Ordem) & ")"
txtrevisao.Text = IIf(IsNull(TBHistProc!Ordem), "", TBHistProc!Ordem)
txtAnalisado.Text = IIf(IsNull(TBHistProc!por), "", (TBHistProc!por))
txtHistorico = IIf(IsNull(TBHistProc!Historico), "", (TBHistProc!Historico))
If IsNull(TBHistProc!Data) = False Then txtDtRevisao.Value = TBHistProc!Data
Novo_Processo_Historico = False

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
    Case 3: ProcAnterior
    Case 4: ProcProximo
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
