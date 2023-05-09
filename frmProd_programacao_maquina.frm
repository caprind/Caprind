VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProd_programacao_maquina 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Programação da produção - Localizar posto de trabalho"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   6060
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   11085
      MouseIcon       =   "frmProd_programacao_maquina.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmProd_programacao_maquina.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela e fecha formulário (Esc)"
      Top             =   165
      Width           =   570
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   6
      Top             =   870
      Width           =   8805
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4620
         TabIndex        =   8
         Top             =   210
         Width           =   3975
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            MouseIcon       =   "frmProd_programacao_maquina.frx":1194
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            MouseIcon       =   "frmProd_programacao_maquina.frx":12E6
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            MouseIcon       =   "frmProd_programacao_maquina.frx":1438
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmProd_programacao_maquina.frx":158A
         Left            =   180
         List            =   "frmProd_programacao_maquina.frx":159A
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   3645
         TabIndex        =   9
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1935
         TabIndex        =   7
         Top             =   180
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView Listamaquina 
      Height          =   3615
      Left            =   60
      TabIndex        =   10
      Top             =   2400
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      MouseIcon       =   "frmProd_programacao_maquina.frx":15C8
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   8140
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Agrega custo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Liberada"
         Object.Width           =   2117
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   885
      Left            =   90
      TabIndex        =   12
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1561
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
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
      ButtonKey1      =   "2"
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
      ButtonWidth1    =   36
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
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   48
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "14"
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
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "15"
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "16"
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
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5850
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProd_programacao_maquina.frx":18E2
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmProd_programacao_maquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Ordem_programacao_maquina As String

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    If cmbfiltrarpor = "Código do posto" Then
        If Optinicio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Maquina like '" & txtTexto & "%' order by maquina"
        If Optmeio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Maquina like '%" & txtTexto & "%' order by maquina"
        If Optfim.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Maquina like '%" & txtTexto & "' order by maquina"
    End If
    If cmbfiltrarpor = "Descrição" Then
        If Optinicio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Descricao like '" & txtTexto & "%' order by maquina"
        If Optmeio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Descricao like '%" & txtTexto & "%' order by maquina"
        If Optfim.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Descricao like '%" & txtTexto & "' order by maquina"
    End If
    If cmbfiltrarpor = "Grupo" Then
        If Optinicio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where grupo like '" & txtTexto & "%' order by maquina"
        If Optmeio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where grupo like '%" & txtTexto & "%' order by maquina"
        If Optfim.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where grupo like '%" & txtTexto & "' order by maquina"
    End If
    If cmbfiltrarpor = "Setor" Then
        If Optinicio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Setor like '" & txtTexto & "%' order by maquina"
        If Optmeio.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Setor like '%" & txtTexto & "%' order by maquina"
        If Optfim.Value = True Then StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas where Setor like '%" & txtTexto & "' order by maquina"
    End If
Else
    StrSql_Ordem_programacao_maquina = "Select * FROM cadmaquinas order by maquina"
End If
ProcAtualizalista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfiltrarpor = "Código do posto"

ProcCarregaToolBar1 Me, 15195, 4, True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista()
On Error GoTo tratar_erro

ListaMaquina.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Ordem_programacao_maquina, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    With ListaMaquina.ListItems
        Do While TBLISTA.EOF = False
            .Add , , ""
            .Item(.Count).SubItems(1) = Trim(TBLISTA!maquina)
            .Item(.Count).SubItems(2) = Trim(TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = Trim(TBLISTA!IDMaquina)
            If TBLISTA!custos = True Then .Item(.Count).SubItems(4) = "SIM" Else .Item(.Count).SubItems(4) = "NÃO"
            If TBLISTA!Liberada = True Then .Item(.Count).SubItems(5) = "SIM" Else .Item(.Count).SubItems(5) = "NÃO"
            TBLISTA.MoveNext
            contador = contador + 1
            PBLista.Value = contador
        Loop
    End With
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaquina_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ListaMaquina.ListItems.Count = 0 Then Exit Sub
ProcOrdenaListView ListaMaquina, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaquina_DblClick()
On Error GoTo tratar_erro

frmProd_programacao.txtIDmaquina = ListaMaquina.SelectedItem.SubItems(3)
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

