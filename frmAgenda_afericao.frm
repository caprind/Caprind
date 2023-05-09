VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAgenda_afericao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contorle de qualidade - Instrumentos - Agenda de calibração"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "frmAgenda_afericao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView Lista 
      Height          =   4890
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8625
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "t"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Cod. ref."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "N. série"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3100
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "D"
         Text            =   "Dt. aquis."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Fabricante"
         Object.Width           =   3100
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "D"
         Text            =   "Dt. calib."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Órgão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "D"
         Text            =   "Próx. calib."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Certificado"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   65
      TabIndex        =   1
      Top             =   990
      Width           =   11805
      Begin VB.TextBox txtdata 
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
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   10905
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data :"
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
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   450
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   1720
      ButtonCount     =   3
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Ajuda"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Ajuda (F1)"
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
      ButtonWidth1    =   41
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Sair"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Sair (Esc)"
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
      ButtonLeft2     =   45
      ButtonTop2      =   2
      ButtonWidth2    =   30
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   5
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   24
      ButtonHeight3   =   24
      ButtonUseMaskColor3=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7380
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmAgenda_afericao.frx":212A
         Count           =   1
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   6540
      Width           =   11805
      _ExtentX        =   20823
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
Attribute VB_Name = "frmAgenda_afericao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11805, 3, True

txtData.Text = Format(Date, "Long Date")
CamposFiltro = "I.CODIGO, I.Numero, EC.Ref, EC.Numero_serie, I.Descricao, I.Data_Aquisicao, I.Fabricante, A.Aferido, A.Orgao, A.Proxima_afericao, A.Certificado"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CamposFiltro & " from (((Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque) LEFT JOIN Projproduto P ON P.Desenho = I.Numero) LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto) INNER JOIN Afericao A ON I.Codigo = A.ID_inst and I.ID_ultima_afericao = A.Codigo where A.proxima_afericao = '" & Format(Date, "Short Date") & "' order by I.numero", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Numero), "", TBLISTA!Numero)
            If IsNull(TBLISTA!Ref) = True Or TBLISTA!Ref = "" Then .Item(.Count).SubItems(2) = FunCarregaCodRef(TBLISTA!Numero) Else .Item(.Count).SubItems(2) = TBLISTA!Ref
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Numero_serie), "", TBLISTA!Numero_serie)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Data_Aquisicao), "", Format(TBLISTA!Data_Aquisicao, "dd/mm/yy"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Fabricante), "", TBLISTA!Fabricante)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Aferido), "", Format(TBLISTA!Aferido, "dd/mm/yy"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Orgao), "", TBLISTA!Orgao)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Proxima_afericao), "", Format(TBLISTA!Proxima_afericao, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
         End With
         TBLISTA.MoveNext
         contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgSair_Click()
On Error GoTo tratar_erro

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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
With frmInstrumentos
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select I.*, EC.Numero_serie, EC.ref from Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.Codigo = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        .ProcLimpar
        .ProcPuxaDados
        .StrSql_Instrumentos_Localizar = "Select I.CODIGO, I.Numero, EC.ref, EC.Numero_serie, I.Descricao, I.Data_Aquisicao, I.Fabricante, I.Familia from Instrumentos I INNER JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.Codigo = " & Lista.SelectedItem
        .ProcCarregaLista (1)
    End If
    TBProduto.Close
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
