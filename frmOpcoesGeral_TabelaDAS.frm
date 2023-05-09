VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpcoesGeral_TabelaDAS 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações do sistema - Opções gerais - Tabela do DAS"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5760
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
   ScaleHeight     =   6900
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alíquotas (%)"
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
      Height          =   885
      Left            =   3390
      TabIndex        =   11
      Top             =   990
      Width           =   2325
      Begin VB.TextBox Txt_ICMS 
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
         Left            =   1170
         MaxLength       =   255
         TabIndex        =   3
         ToolTipText     =   "Alíquota do ICMS."
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox Txt_DAS 
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
         MaxLength       =   255
         TabIndex        =   2
         ToolTipText     =   "Alíquota do DAS."
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1470
         TabIndex        =   13
         Top             =   210
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DAS"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   517
         TabIndex        =   12
         Top             =   210
         Width           =   300
      End
   End
   Begin VB.TextBox Txt_ID 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   5
      Text            =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   950
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4725
      Left            =   60
      TabIndex        =   4
      Top             =   1890
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "De"
         Object.Width           =   2615
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Até"
         Object.Width           =   2615
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "DAS (%)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "ICMS (%)"
         Object.Width           =   1764
      EndProperty
   End
   Begin DrawSuite2014.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   6630
      Width           =   5655
      _ExtentX        =   9975
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
   Begin DrawSuite2014.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
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
      Begin DrawSuite2014.USImageList USImageList1 
         Left            =   3720
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmOpcoesGeral_TabelaDAS.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Receita bruta em 12 meses (em R$)"
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
      Height          =   885
      Left            =   55
      TabIndex        =   6
      Top             =   990
      Width           =   3315
      Begin VB.TextBox Txt_ate 
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
         Left            =   1680
         MaxLength       =   255
         TabIndex        =   1
         ToolTipText     =   "Valor até."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox Txt_de 
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
         MaxLength       =   255
         TabIndex        =   0
         ToolTipText     =   "Valor de."
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   825
         TabIndex        =   7
         Top             =   210
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmOpcoesGeral_TabelaDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Tabela_DAS As Boolean 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente excluir este(s) registro(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            
            Permitido = True
            Conexao.Execute "DELETE from Impostos_TabelaDAS where ID = " & .ListItems(InitFor)
            '==================================
            Modulo = "Configuração do sistema/Opções gerais"
            Evento = "Excluir registro da tabela do DAS"
            ID_documento = .ListItems(InitFor)
            Documento = "Empresa: " & frmOpcoesGeral.txtRazao
            Documento1 = "De: " & .ListItems(InitFor).SubItems(1) & " - Até: " & .ListItems(InitFor).SubItems(2) & " - DAS: " & .ListItems(InitFor).SubItems(3) & " - ICMS: " & .ListItems(InitFor).SubItems(4)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) registro(s) antes de excluir."), vbExclamation
Else
    MsgBox ("Registro(s) excluído(s) com sucesso."), vbInformation
    ProcLimpaCampos
    Frame1.Enabled = False
    Frame2.Enabled = False
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
Valor = IIf(Txt_de = "", 0, Txt_de)
If Txt_de = "" Or Valor < 0 Then
    NomeCampo = "o valor de"
    ProcVerificaAcao
    Txt_de.SetFocus
    Exit Sub
End If
Valor = IIf(Txt_ate = "", 0, Txt_ate)
If Txt_ate = "" Or Valor < 0 Then
    NomeCampo = "o valor até"
    ProcVerificaAcao
    Txt_ate.SetFocus
    Exit Sub
End If
Valor = IIf(Txt_DAS = "", 0, Txt_DAS)
If Valor <= 0 Then
    NomeCampo = "a alíquota do DAS"
    ProcVerificaAcao
    Txt_DAS.SetFocus
    Exit Sub
End If
Valor = IIf(Txt_ICMS = "", 0, Txt_ICMS)
If Valor <= 0 Then
    NomeCampo = "a alíquota do ICMS"
    ProcVerificaAcao
    Txt_ICMS.SetFocus
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Impostos_TabelaDAS where ID = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_empresa = frmOpcoesGeral.txtidempresa
TBGravar!De = Txt_de
TBGravar!Ate = Txt_ate
TBGravar!DAS = Txt_DAS
TBGravar!ICMS = Txt_ICMS
TBGravar.Update
Txt_ID = TBGravar!ID
TBGravar.Close
ProcCarregaLista
If Novo_Tabela_DAS = True Then
    MsgBox ("Novo registro cadastrado com sucesso."), vbInformation
    Evento = "Novo registro da tabela do DAS"
Else
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar registro da tabela do DAS"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
Modulo = "Configuração do sistema/Opções gerais"
ID_documento = Txt_ID
Documento = "Empresa: " & frmOpcoesGeral.txtRazao
Documento1 = "De: " & Format(Txt_de, "###,##0.00") & " - Até: " & Format(Txt_ate, "###,##0.00") & " - DAS: " & Format(Txt_DAS, "###,##0.00") & " - ICMS: " & Format(Txt_ICMS, "###,##0.00")
ProcGravaEvento
'==================================
Novo_Tabela_DAS = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Tabela_DAS = True
Frame2.Enabled = True
Frame1.Enabled = True
Txt_de.SetFocus

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6165, 7, True
ProcLimpaVariaveisPrincipais
ProcCarregaLista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Tabela_DAS = True Then
    If MsgBox("O registro ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo + vbQuestion) = vbYes Then
        ProcSalvar
        If Novo_Tabela_DAS = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Tabela_DAS = False
Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Impostos_TabelaDAS where ID_empresa = " & frmOpcoesGeral.txtidempresa, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!De), "", Format(TBLISTA!De, "###,##0.00"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ate), "", Format(TBLISTA!Ate, "###,##0.00"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!DAS), "", Format(TBLISTA!DAS, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!ICMS), "", Format(TBLISTA!ICMS, "###,##0.00"))
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
Txt_de = ""
Txt_ate = ""
Txt_DAS = ""
Txt_ICMS = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Impostos_TabelaDAS where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    CodigoLista = Lista.SelectedItem.Index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

Txt_ID = TBLISTA!ID
Txt_de = IIf(IsNull(TBLISTA!De), "", Format(TBLISTA!De, "###,##0.00"))
Txt_ate = IIf(IsNull(TBLISTA!Ate), "", Format(TBLISTA!Ate, "###,##0.00"))
Txt_DAS = IIf(IsNull(TBLISTA!DAS), "", Format(TBLISTA!DAS, "###,##0.00"))
Txt_ICMS = IIf(IsNull(TBLISTA!ICMS), "", Format(TBLISTA!ICMS, "###,##0.00"))
Frame1.Enabled = True
Frame2.Enabled = True
Novo_Tabela_DAS = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_DAS_Change()
On Error GoTo tratar_erro
    
If Txt_DAS <> "" Then
    VerifNumero = Txt_DAS
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_DAS = ""
        Txt_DAS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_DAS_LostFocus()
On Error GoTo tratar_erro

Txt_DAS = Format(Txt_DAS, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_ate_Change()
On Error GoTo tratar_erro
    
If Txt_ate <> "" Then
    VerifNumero = Txt_ate
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ate = ""
        Txt_ate.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_ate_LostFocus()
On Error GoTo tratar_erro

Txt_ate = Format(Txt_ate, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_De_Change()
On Error GoTo tratar_erro
    
If Txt_de <> "" Then
    VerifNumero = Txt_de
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_de = ""
        Txt_de.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_de_LostFocus()
On Error GoTo tratar_erro

Txt_de = Format(Txt_de, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_ICMS_Change()
On Error GoTo tratar_erro
    
If Txt_ICMS <> "" Then
    VerifNumero = Txt_ICMS
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ICMS = ""
        Txt_ICMS.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_ICMS_LostFocus()
On Error GoTo tratar_erro

Txt_ICMS = Format(Txt_ICMS, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
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
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
