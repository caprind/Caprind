VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmContas_antecipacoes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Contas pagas - Lista de antecipações"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
   ClipControls    =   0   'False
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
   ScaleHeight     =   6540
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_valor_antecipacao 
      Alignment       =   1  'Right Justify
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
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Valor da antecipação."
      Top             =   1080
      Width           =   1620
   End
   Begin VB.TextBox txtTotalContas 
      Alignment       =   1  'Right Justify
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
      Left            =   6285
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Saldo."
      Top             =   6180
      Width           =   1620
   End
   Begin DrawSuite2022.USProgressBar PbLista 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   6210
      Width           =   5475
      _ExtentX        =   9657
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5310
      Top             =   2040
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmContas_antecipacoes.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
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
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4665
      Left            =   60
      TabIndex        =   1
      Top             =   1470
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   8229
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "IDconta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Dt. vencto."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Nº documento"
         Object.Width           =   5706
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor da antecipação :"
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
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   1080
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo :"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   6180
      Width           =   615
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmContas_antecipacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7875, 3, True

Label1.Visible = False
Txt_valor_antecipacao.Visible = False
With Lista
    .Top = 1080
    .Height = 4665
End With
With PBLista
    .Top = 5760
    .Width = 7875
End With
lblTotal.Visible = False
txtTotalContas.Visible = False
Height = 6420

If Financeiro_Contas_Pagas = True Then Caption = "Financeiro - Contas pagas - Lista de antecipações"
If Financeiro_Contas_Recebidas = True Then Caption = "Financeiro - Contas recebidas - Lista de antecipações"
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Receber = True Then
    If Financeiro_Contas_Pagar = True Then
        Caption = "Financeiro - Contas a pagar - Lista de contas relacionadas"
        With frmContas_Pagar
            'Valor da antecipação
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select dbl_valorpagto from tbl_contaspagar where idintconta = " & .txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtd = IIf(IsNull(TBAbrir!dbl_valorpagto), 0, TBAbrir!dbl_valorpagto)
            End If
            Txt_valor_antecipacao = Format(Qtd, "###,##0.00")
            
            'Valor utilizado
            Qtde = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "SELECT Sum(Valor) as valor from tbl_Contas_antecipacao where id_antecipacao = " & .txtidintconta & " and tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            TBAbrir.Close
            
            qt = Qtd - Qtde
            txtTotalContas = IIf(qt < 0, 0, Format(qt, "###,##0.00"))
        End With
    End If
    If Financeiro_Contas_Receber = True Then
        Caption = "Financeiro - Contas a receber - Lista de contas relacionadas"
        With frmContas_Receber
            'Valor da antecipação
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Valor from tbl_contas_receber where idintconta = " & .txtidintconta, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtd = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            Txt_valor_antecipacao = Format(Qtd, "###,##0.00")
            
            'Valor utilizado
            Qtde = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "SELECT Sum(Valor) as valor from tbl_Contas_antecipacao where id_antecipacao = " & .txtidintconta & " and tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Qtde = IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
            End If
            TBAbrir.Close
            
            qt = Qtd - Qtde
            txtTotalContas = IIf(qt < 0, 0, Format(qt, "###,##0.00"))
        End With
    End If
    
    Label1.Visible = True
    Txt_valor_antecipacao.Visible = True
    With Lista
        .Top = 1470
        .Height = 4665
    End With
    With PBLista
        .Top = 6210
        .Width = 5475
    End With
    lblTotal.Visible = True
    txtTotalContas.Visible = True
    Height = 6900
End If
ProcCarregaListaAntecipacao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaAntecipacao()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas Then
    INNERJOINTEXTO = "Select CA.ID, CA.valor, CP.* from tbl_ContasPagar CP INNER JOIN tbl_contas_antecipacao CA ON CA.ID_conta = CP.IDintconta"
    If Financeiro_Contas_Pagar = True Then
        TextoFiltro = "CA.ID_antecipacao = " & frmContas_Pagar.txtidintconta
    Else
        With frmContas_Pagas
            If .txtStatus = "TÍTULO LIQUIDADO ANTECIPADO" Then TextoFiltro = "CA.ID_antecipacao = " & .txtidintconta Else TextoFiltro = "CA.ID_conta = " & .txtidintconta
        End With
    End If
    TBLISTA.Open INNERJOINTEXTO & " where " & TextoFiltro & " and CA.Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
Else
    INNERJOINTEXTO = "Select CA.ID, CA.valor, CR.* from tbl_contas_receber CR INNER JOIN tbl_contas_antecipacao CA ON CA.ID_conta = CR.IDintconta"
    If Financeiro_Contas_Receber = True Then
        TextoFiltro = "CA.ID_antecipacao = " & frmContas_Receber.txtidintconta
    Else
        With frmContas_recebidas
            If .txtStatus = "TÍTULO LIQUIDADO ANTECIPADO" Then TextoFiltro = "CA.ID_antecipacao = " & .txtidintconta Else TextoFiltro = "CA.ID_conta = " & .txtidintconta
        End With
    End If
    TBLISTA.Open INNERJOINTEXTO & " where " & TextoFiltro & " and CA.Tipo = 'R'", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!IDintconta
            .Item(.Count).SubItems(4) = Format(TBLISTA!valor, "###,##0.00")
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!txt_ndocumento), "", TBLISTA!txt_ndocumento)
            If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Pagas = True Then
                DataEmissao = TBLISTA!Dt_emissao
                DataVencimento = TBLISTA!dt_Pagamento
                Parcela = TBLISTA!txt_Parcela
            Else
                DataEmissao = TBLISTA!emissao
                DataVencimento = TBLISTA!Vencimento
                Parcela = TBLISTA!Parcela
            End If
            .Item(.Count).SubItems(2) = Format(DataEmissao, "dd/mm/yy")
            .Item(.Count).SubItems(3) = Format(DataVencimento, "dd/mm/yy")
            .Item(.Count).SubItems(6) = Parcela
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
