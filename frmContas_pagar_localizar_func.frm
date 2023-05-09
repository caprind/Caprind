VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmContas_pagar_localizar_func 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Contas à pagar - Localizar funcionário"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9405
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
   ScaleHeight     =   7005
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6840
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmContas_pagar_localizar_func.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   9
      Top             =   990
      Width           =   9285
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4290
         TabIndex        =   14
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            Height          =   255
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmContas_pagar_localizar_func.frx":21EE
         Left            =   180
         List            =   "frmContas_pagar_localizar_func.frx":2207
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4035
      End
      Begin VB.ComboBox cmbTexto 
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
         Height          =   330
         ItemData        =   "frmContas_pagar_localizar_func.frx":2235
         Left            =   180
         List            =   "frmContas_pagar_localizar_func.frx":2237
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   8895
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Número do CPF."
         Top             =   1050
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
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
         Width           =   8895
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
         Left            =   3892
         TabIndex        =   11
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
         Left            =   1777
         TabIndex        =   10
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   12
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
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
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
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
      ButtonWidth1    =   42
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
      ButtonLeft2     =   46
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
      ButtonLeft3     =   50
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
      ButtonLeft4     =   93
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   6720
      Width           =   9285
      _ExtentX        =   16378
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
   Begin MSComctlLib.ListView Lista 
      Height          =   4185
      Left            =   60
      TabIndex        =   4
      Top             =   2520
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   7382
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nome"
         Object.Width           =   13573
      EndProperty
   End
End
Attribute VB_Name = "frmContas_pagar_localizar_func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Contas_Localizar_Funcionarios As String 'OK

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
With cmbTexto
    .Clear
    If cmbfiltrarpor = "Setor" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select setor from Usuarios_Setor order by setor", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                .AddItem TBAbrir!Setor
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    ElseIf cmbfiltrarpor = "Sexo" Then
            .AddItem "Feminino"
            .AddItem "Masculino"
        ElseIf cmbfiltrarpor = "Tipo" Then
                .AddItem "Horista"
                .AddItem "Mensalista"
            ElseIf cmbfiltrarpor = "CPF" Then
                    txtTexto.Visible = False
                    .Visible = False
                    txtCpf.Visible = True
                    Exit Sub
                Else
                    txtCpf.Visible = False
                    txtTexto.Visible = True
                    .Visible = False
                    Exit Sub
    End If
    txtTexto.Visible = False
    .Visible = True
    txtCpf.Visible = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If cmbTexto <> "" Then
    txtTexto = ""
    txtCpf = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

INNERJOINTEXTO = "Select ID, Codigo, Nome from Funcionarios"
If Financeiro_Contas_Pagar = True Or Financeiro_Contas_Receber = True Then TextoFiltroPadrao = "DtValidacao IS NOT NULL and Situacao = 'Normal' order by nome" Else TextoFiltroPadrao = " order by nome"

If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Then
    If cmbfiltrarpor = "Sexo" Or cmbfiltrarpor = "Tipo" Or cmbfiltrarpor = "Setor" Then
        Select Case cmbfiltrarpor
            Case "Sexo": TextoFiltro = "Sexo"
            Case "Tipo": TextoFiltro = "tipo"
            Case "Setor": TextoFiltro = "setor"
        End Select
        StrSql_Contas_Localizar_Funcionarios = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "CPF" Then
            StrSql_Contas_Localizar_Funcionarios = INNERJOINTEXTO & " where CPF = '" & txtCpf.Text & "' and " & TextoFiltroPadrao
        Else
            Select Case cmbfiltrarpor
                Case "Código": TextoFiltro = "Codigo"
                Case "Nome": TextoFiltro = "Nome"
                Case "RG": TextoFiltro = "RG"
            End Select
            StrSql_Contas_Localizar_Funcionarios = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Contas_Localizar_Funcionarios = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If
ProcAtualizalista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista()
On Error GoTo tratar_erro

If StrSql_Contas_Localizar_Funcionarios = "" Then Exit Sub
Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Contas_Localizar_Funcionarios, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Nome), "", Trim(TBLISTA!Nome))
        End With
        contador = contador + 1
        PBLista.Value = contador
        TBLISTA.MoveNext
    Loop
    
End If
TBLISTA.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: Lista_DblClick
    Case vbKeyF2: ProcFiltrar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9285, 5, True

If Financeiro_Contas_Pagas = True Then Caption = "Administrativo - Financeiro - Contas pagas - Localizar funcionário"
If Financeiro_Contas_Receber = True Then Caption = "Administrativo - Financeiro - Contas à receber - Localizar funcionário"
If Financeiro_Contas_Recebidas = True Then Caption = "Administrativo - Financeiro - Contas recebidas - Localizar funcionário"

cmbfiltrarpor = "Nome"
txtTexto.Visible = True
cmbTexto.Visible = False
txtCpf.Visible = False

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
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * FROM Funcionarios WHERE ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If Financeiro_Contas_Pagar = True Then
        With frmContas_Pagar
            .Cmb_tipo = "Funcionário"
            .txtIDFornec = TBFornecedor!CODIGO
        End With
    ElseIf Financeiro_Contas_Pagas = True Then
            With frmContas_Pagas
                .Cmb_tipo = "Funcionário"
                .txtIDFornec = TBFornecedor!CODIGO
            End With
        ElseIf Financeiro_Contas_Receber = True Then
                With frmContas_Receber
                    .Cmb_tipo = "Funcionário"
                    .txtIDcliente = TBFornecedor!CODIGO
                End With
            Else
                With frmContas_recebidas
                    .Cmb_tipo = "Funcionário"
                    .txtIDcliente = TBFornecedor!CODIGO
                End With
    End If
End If
TBFornecedor.Close
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCpf_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If txtCpf <> "___.___.___-__" Then
    cmbTexto.Clear
    txtTexto = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
If txtTexto <> "" Then
    cmbTexto.Clear
    txtCpf = "___.___.___-__"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
