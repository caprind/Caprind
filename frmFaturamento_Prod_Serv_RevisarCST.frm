VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_RevisarCST 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Revisar CST"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CST (cadastrada)"
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
      Height          =   885
      Left            =   60
      TabIndex        =   11
      Top             =   1020
      Width           =   3285
      Begin VB.TextBox Txt_CST_Cofins 
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
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Situação tributária Cofins."
         Top             =   420
         Width           =   660
      End
      Begin VB.TextBox Txt_CST_PIS 
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
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Situação tributária PIS."
         Top             =   420
         Width           =   650
      End
      Begin VB.TextBox Txt_CST_IPI 
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
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Situação tributária IPI."
         Top             =   420
         Width           =   650
      End
      Begin VB.TextBox Txt_CST_ICMS 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Situação tributária ICMS."
         Top             =   420
         Width           =   920
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS           IPI         PIS       Cofins"
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
         Left            =   465
         TabIndex        =   12
         Top             =   230
         Width           =   2490
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2310
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_Prod_Serv_RevisarCST.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CST (permitida)"
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
      Height          =   885
      Left            =   60
      TabIndex        =   9
      Top             =   1920
      Width           =   3285
      Begin VB.ComboBox txtCST_Cofins 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_RevisarCST.frx":1E05
         Left            =   2430
         List            =   "frmFaturamento_Prod_Serv_RevisarCST.frx":1E6C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Situação tributária Cofins."
         Top             =   420
         Width           =   660
      End
      Begin VB.ComboBox txtCST_PIS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_RevisarCST.frx":1EF2
         Left            =   1770
         List            =   "frmFaturamento_Prod_Serv_RevisarCST.frx":1F59
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Situação tributária PIS."
         Top             =   420
         Width           =   660
      End
      Begin VB.ComboBox txtCST_IPI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_RevisarCST.frx":1FDF
         Left            =   1110
         List            =   "frmFaturamento_Prod_Serv_RevisarCST.frx":2010
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Situação tributária IPI."
         Top             =   420
         Width           =   660
      End
      Begin VB.ComboBox txtCST_ICMS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmFaturamento_Prod_Serv_RevisarCST.frx":204E
         Left            =   180
         List            =   "frmFaturamento_Prod_Serv_RevisarCST.frx":2109
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Situação tributária ICMS."
         Top             =   420
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS           IPI         PIS       Cofins"
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
         Index           =   10
         Left            =   465
         TabIndex        =   10
         Top             =   230
         Width           =   2490
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_RevisarCST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar a CST"
If txtCST_ICMS.Text = "" Then
    NomeCampo = "do ICMS"
    ProcVerificaAcao
    txtCST_ICMS.SetFocus
    Exit Sub
End If
If txtCST_IPI.Text = "" Then
    NomeCampo = "do IPI"
    ProcVerificaAcao
    txtCST_IPI.SetFocus
    Exit Sub
End If
If txtCST_PIS.Text = "" Then
    NomeCampo = "do PIS"
    ProcVerificaAcao
    txtCST_PIS.SetFocus
    Exit Sub
End If
If txtCST_Cofins.Text = "" Then
    NomeCampo = "do Cofins"
    ProcVerificaAcao
    txtCST_Cofins.SetFocus
    Exit Sub
End If

With frmFaturamento_Prod_Serv
    Set TBOS = CreateObject("adodb.recordset")
    TBOS.Open "Select * from tbl_detalhes_nota where Int_codigo = " & .txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBOS.EOF = False Then
        TBOS!txt_CST = txtCST_ICMS.Text
        TBOS!CST_IPI = txtCST_IPI.Text
        TBOS!CST_PIS = txtCST_PIS.Text
        TBOS!CST_Cofins = txtCST_Cofins.Text
        TBOS.Update
    End If
    TBOS.Close
    
    '==================================
    Modulo = Formulario
    Evento = "Revisar CST"
    ID_documento = .txtidproduto
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtID Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    Documento1 = "Cód. interno: " & .txtCod_Produto
    ProcGravaEvento
    '==================================
    
    .txtCST_ICMS = txtCST_ICMS
    .txtCST_IPI = txtCST_IPI
    .txtCST_PIS = txtCST_PIS
    .txtCST_Cofins = txtCST_Cofins
End With
USMsgBox ("CST revisada com sucesso."), vbExclamation, "CAPRIND v5.0"

Unload Me

Exit Sub
tratar_erro:
    If Err.Number = "383" Then Unload Me
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
    
Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 3255, 5, True
ProcCarregaCST

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

Private Sub ProcCarregaCST()
On Error GoTo tratar_erro

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select txt_CST, CST_IPI, CST_PIS, CST_Cofins from tbl_Detalhes_Nota where Int_codigo = " & frmFaturamento_Prod_Serv.txtidproduto, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    Txt_CST_ICMS = IIf(IsNull(TBCFOP!txt_CST), "", TBCFOP!txt_CST)
    Txt_CST_IPI = IIf(IsNull(TBCFOP!CST_IPI), "", TBCFOP!CST_IPI)
    Txt_CST_PIS = TBCFOP!CST_PIS
    Txt_CST_Cofins = TBCFOP!CST_Cofins
End If
TBCFOP.Close

txtCST_ICMS.Clear
txtCST_IPI.Clear
txtCST_PIS.Clear
txtCST_Cofins.Clear
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select ID from tbl_NaturezaOperacao_CST where ID_CFOP = " & frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
    'CST de ICMS
    contador = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CST_ICMS from tbl_NaturezaOperacao_CST where ID_CFOP = " & frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod & " group by CST_ICMS", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            If IsNull(TBCFOP!CST_ICMS) = False And TBCFOP!CST_ICMS <> "" Then
                txtCST_ICMS.AddItem TBCFOP!CST_ICMS
                contador = contador + 1
                CSTTexto = TBCFOP!CST_ICMS
            End If
            TBCFOP.MoveNext
        Loop
        If contador = 1 Then txtCST_ICMS = CSTTexto
    End If
    
    'CST de IPI
    contador = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CST_IPI from tbl_NaturezaOperacao_CST where ID_CFOP = " & frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod & " group by CST_IPI", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            If IsNull(TBCFOP!CST_IPI) = False And TBCFOP!CST_IPI <> "" Then
                txtCST_IPI.AddItem TBCFOP!CST_IPI
                contador = contador + 1
                CSTTexto = TBCFOP!CST_IPI
            End If
            TBCFOP.MoveNext
        Loop
        If contador = 1 Then txtCST_IPI = CSTTexto
    End If
    
    'CST de PIS
    contador = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CST_PIS from tbl_NaturezaOperacao_CST where ID_CFOP = " & frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod & " group by CST_PIS", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            If IsNull(TBCFOP!CST_PIS) = False And TBCFOP!CST_PIS <> "" Then
                txtCST_PIS.AddItem TBCFOP!CST_PIS
                contador = contador + 1
                CSTTexto = TBCFOP!CST_PIS
            End If
            TBCFOP.MoveNext
        Loop
        If contador = 1 Then txtCST_PIS = CSTTexto
    End If
    
    'CST de Cofins
    contador = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CST_Cofins from tbl_NaturezaOperacao_CST where ID_CFOP = " & frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod & " group by CST_Cofins", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Do While TBCFOP.EOF = False
            If IsNull(TBCFOP!CST_Cofins) = False And TBCFOP!CST_Cofins <> "" Then
                txtCST_Cofins.AddItem TBCFOP!CST_Cofins
                contador = contador + 1
                CSTTexto = TBCFOP!CST_Cofins
            End If
            TBCFOP.MoveNext
        Loop
        If contador = 1 Then txtCST_Cofins = CSTTexto
    End If
Else
    With txtCST_ICMS
        .AddItem "000"
        .AddItem "010"
        .AddItem "0101"
        .AddItem "0102"
        .AddItem "0103"
        .AddItem "020"
        .AddItem "0201"
        .AddItem "0202"
        .AddItem "0203"
        .AddItem "0300"
        .AddItem "040"
        .AddItem "0400"
        .AddItem "041"
        .AddItem "050"
        .AddItem "0500"
        .AddItem "051"
        .AddItem "060"
        .AddItem "070"
        .AddItem "090"
        .AddItem "0900"
        
        .AddItem "100"
        .AddItem "110"
        .AddItem "1101"
        .AddItem "1102"
        .AddItem "1103"
        .AddItem "120"
        .AddItem "1201"
        .AddItem "1202"
        .AddItem "1203"
        .AddItem "1300"
        .AddItem "140"
        .AddItem "1400"
        .AddItem "141"
        .AddItem "150"
        .AddItem "1500"
        .AddItem "151"
        .AddItem "160"
        .AddItem "170"
        .AddItem "190"
        .AddItem "1900"
        
        .AddItem "200"
        .AddItem "210"
        .AddItem "2101"
        .AddItem "2102"
        .AddItem "2103"
        .AddItem "220"
        .AddItem "2201"
        .AddItem "2202"
        .AddItem "2203"
        .AddItem "2300"
        .AddItem "240"
        .AddItem "2400"
        .AddItem "241"
        .AddItem "250"
        .AddItem "2500"
        .AddItem "251"
        .AddItem "260"
        .AddItem "270"
        .AddItem "290"
        .AddItem "2900"
        
        .AddItem "300"
        .AddItem "310"
        .AddItem "3101"
        .AddItem "3102"
        .AddItem "3103"
        .AddItem "320"
        .AddItem "3201"
        .AddItem "3202"
        .AddItem "3203"
        .AddItem "3300"
        .AddItem "340"
        .AddItem "3400"
        .AddItem "341"
        .AddItem "350"
        .AddItem "3500"
        .AddItem "351"
        .AddItem "360"
        .AddItem "370"
        .AddItem "390"
        .AddItem "3900"
        
        .AddItem "400"
        .AddItem "410"
        .AddItem "4101"
        .AddItem "4102"
        .AddItem "4103"
        .AddItem "420"
        .AddItem "4201"
        .AddItem "4202"
        .AddItem "4203"
        .AddItem "4300"
        .AddItem "440"
        .AddItem "4400"
        .AddItem "441"
        .AddItem "450"
        .AddItem "4500"
        .AddItem "451"
        .AddItem "460"
        .AddItem "470"
        .AddItem "490"
        .AddItem "4900"
        
        .AddItem "500"
        .AddItem "510"
        .AddItem "5101"
        .AddItem "5102"
        .AddItem "5103"
        .AddItem "520"
        .AddItem "5201"
        .AddItem "5202"
        .AddItem "5203"
        .AddItem "5300"
        .AddItem "540"
        .AddItem "5400"
        .AddItem "541"
        .AddItem "550"
        .AddItem "5500"
        .AddItem "551"
        .AddItem "560"
        .AddItem "570"
        .AddItem "590"
        .AddItem "5900"
        
        .AddItem "600"
        .AddItem "610"
        .AddItem "6101"
        .AddItem "6102"
        .AddItem "6103"
        .AddItem "620"
        .AddItem "6201"
        .AddItem "6202"
        .AddItem "6203"
        .AddItem "6300"
        .AddItem "640"
        .AddItem "6400"
        .AddItem "641"
        .AddItem "650"
        .AddItem "6500"
        .AddItem "651"
        .AddItem "660"
        .AddItem "670"
        .AddItem "690"
        .AddItem "6900"
        
        .AddItem "700"
        .AddItem "710"
        .AddItem "7101"
        .AddItem "7102"
        .AddItem "7103"
        .AddItem "720"
        .AddItem "7201"
        .AddItem "7202"
        .AddItem "7203"
        .AddItem "7300"
        .AddItem "740"
        .AddItem "7400"
        .AddItem "741"
        .AddItem "750"
        .AddItem "7500"
        .AddItem "751"
        .AddItem "760"
        .AddItem "770"
        .AddItem "790"
        .AddItem "7900"
    End With
    
    With txtCST_IPI
        .AddItem "00"
        .AddItem "01"
        .AddItem "02"
        .AddItem "03"
        .AddItem "04"
        .AddItem "05"
        .AddItem "49"
        .AddItem "50"
        .AddItem "51"
        .AddItem "52"
        .AddItem "53"
        .AddItem "54"
        .AddItem "55"
        .AddItem "99"
    End With

    With txtCST_PIS
        .AddItem "01"
        .AddItem "02"
        .AddItem "03"
        .AddItem "04"
        .AddItem "06"
        .AddItem "07"
        .AddItem "08"
        .AddItem "09"
        .AddItem "49"
        .AddItem "98"
        .AddItem "99"
    End With
    
    With txtCST_Cofins
        .AddItem "01"
        .AddItem "02"
        .AddItem "03"
        .AddItem "04"
        .AddItem "06"
        .AddItem "07"
        .AddItem "08"
        .AddItem "09"
        .AddItem "49"
        .AddItem "98"
        .AddItem "99"
    End With
End If
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
