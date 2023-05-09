VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmManutencao_Agenda 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manutenção - Equipamentos - Agenda de solicitações"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8430
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
   Icon            =   "frmManutencao_Agenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2790
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmManutencao_Agenda.frx":000C
      Count           =   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   55
      TabIndex        =   2
      Top             =   975
      Width           =   8325
      Begin MSComCtl2.DTPicker txtData 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Data."
         Top             =   210
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   490864640
         CurrentDate     =   39057
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4290
      Left            =   60
      TabIndex        =   1
      Top             =   1635
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   7567
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. equip."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3942
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Dt. solic."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Requisitante"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Setor"
         Object.Width           =   2999
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   5940
      Width           =   8325
      _ExtentX        =   14684
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   1720
      ButtonCount     =   4
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonAlignment1=   2
      ButtonType1     =   1
      ButtonStyle1    =   -1
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState1    =   -1
      ButtonLeft1     =   2
      ButtonTop1      =   4
      ButtonWidth1    =   2
      ButtonHeight1   =   54
      ButtonCaption2  =   "Ajuda"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Ajuda (F1)"
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
      ButtonLeft2     =   6
      ButtonTop2      =   2
      ButtonWidth2    =   36
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Sair"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Sair (Esc)"
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
      ButtonLeft3     =   44
      ButtonTop3      =   2
      ButtonWidth3    =   26
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
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
      ButtonState4    =   5
      ButtonLeft4     =   72
      ButtonTop4      =   2
      ButtonWidth4    =   24
      ButtonHeight4   =   24
      ButtonUseMaskColor4=   0   'False
   End
End
Attribute VB_Name = "frmManutencao_Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn: Lista_DblClick
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8520, 4, True
txtData.Value = Date
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdata_Change()
On Error GoTo tratar_erro

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 2: ProcAjuda
    Case 3: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBSolicitacao = CreateObject("adodb.recordset")
TBSolicitacao.Open "Select * from manutencao where Data_Solicitacao = '" & txtData.Value & "' and tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
If TBSolicitacao.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBSolicitacao.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBSolicitacao.EOF = False
        With Lista.ListItems
            .Add , , TBSolicitacao!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBSolicitacao!IDMaquina), "", TBSolicitacao!IDMaquina)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBSolicitacao!Descricao), "", TBSolicitacao!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBSolicitacao!Data_Solicitacao), "", Format(TBSolicitacao!Data_Solicitacao, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBSolicitacao!Requisitante), "", TBSolicitacao!Requisitante)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBSolicitacao!setor_requisitante), "", TBSolicitacao!setor_requisitante)
        End With
        TBSolicitacao.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBSolicitacao.Close

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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
With frmManutencao
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from manutencao where Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        ProcPuxaDados
    End If
    TBAbrir.Close
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

With frmManutencao
    .txttipo.Text = "Corretiva"
    .txtRequisitante.Text = IIf(IsNull(TBAbrir!Requisitante), "", TBAbrir!Requisitante)
    .cmbSetor_Requisitante.Text = IIf(IsNull(TBAbrir!setor_requisitante), "", TBAbrir!setor_requisitante)
    .txtData_Solicitacao.Text = IIf(IsNull(TBAbrir!Data_Solicitacao), "", Format(TBAbrir!Data_Solicitacao, "dd/mm/yyyy"))
    .txtAprovado.Text = IIf(IsNull(TBAbrir!Aprovado), "", TBAbrir!Aprovado)
    .txtSetor_Aprovado.Text = IIf(IsNull(TBAbrir!Setor_Aprovado), "", TBAbrir!Setor_Aprovado)
    .txtLista.Text = IIf(IsNull(TBAbrir!Lista), "", TBAbrir!Lista)
    .txtID = TBAbrir!CODIGO
    .txtIDmaquina = IIf(IsNull(TBAbrir!IDMaquina), "", TBAbrir!IDMaquina)
    .txtdescricao = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
    .txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
    .txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
    If TBAbrir!Controlada = True Then .chkControlada.Value = 1 Else .chkControlada.Value = 0
    .ProcHabilitarPrevCorr
    .Frame2.Enabled = True
    .Frame6.Enabled = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
