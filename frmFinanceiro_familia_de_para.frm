VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFinanceiro_familia_de_para 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Plano de contas - De, para"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6315
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
   ScaleHeight     =   2940
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   8
      Top             =   990
      Width           =   6225
      Begin VB.OptionButton Opt_pagar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contas a pagar"
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
         Left            =   180
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton Opt_receber 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contas a receber"
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
         Left            =   1650
         TabIndex        =   1
         Top             =   210
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1425
      Left            =   55
      TabIndex        =   5
      Top             =   1500
      Width           =   6225
      Begin VB.ComboBox Cmb_para 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Para."
         Top             =   960
         Width           =   5865
      End
      Begin VB.ComboBox Cmb_de 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "De."
         Top             =   390
         Width           =   5865
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para"
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
         Left            =   2947
         TabIndex        =   7
         Top             =   750
         Width           =   330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
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
         Left            =   3015
         TabIndex        =   6
         Top             =   180
         Width           =   195
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3150
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFinanceiro_familia_de_para.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
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
      ButtonKey3      =   "6"
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
      ButtonKey4      =   "7"
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
      ButtonKey5      =   "8"
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
   End
End
Attribute VB_Name = "frmFinanceiro_familia_de_para"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcGravar()
On Error GoTo tratar_erro

Acao = "salvar"
If Cmb_de = "" Then
    NomeCampo = "a subfamília"
    ProcVerificaAcao
    Cmb_de.SetFocus
    Exit Sub
End If
If Cmb_para = "" Then
    NomeCampo = "a classe contábil"
    ProcVerificaAcao
    Cmb_de.SetFocus
    Exit Sub
End If
If Permitido = False Then
    TextoFiltro = "SubFamilia = '" & Cmb_de & "'"
    TextoFiltro1 = "Subfamilia_financeiro = '" & Cmb_de & "'"
Else
    TextoFiltro = "ID_PC = " & Cmb_de.ItemData(Cmb_de.ListIndex)
    TextoFiltro1 = "ID_PC = " & Cmb_de.ItemData(Cmb_de.ListIndex)
End If
Conexao.Execute "Update Familia_financeiro Set ID_PC = " & Cmb_para.ItemData(Cmb_para.ListIndex) & " where " & TextoFiltro
Conexao.Execute "Update projproduto Set ID_PC = " & Cmb_para.ItemData(Cmb_para.ListIndex) & " where " & TextoFiltro1
Conexao.Execute "Update Projfamilia Set ID_PC = " & Cmb_para.ItemData(Cmb_para.ListIndex) & " where " & TextoFiltro1

USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Financeiro/Plano de contas/De, para"
Evento = "Alterar"
ID_documento = 0
Documento = "De: " & Cmb_de & " - Para: " & Cmb_para
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaComboDe

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcGravar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6255, 5, True

ProcCarregaComboDe
ProcCarregaComboPara
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboDe()
On Error GoTo tratar_erro

Permitido = False
With Cmb_de
    .Clear
    If Opt_pagar.Value = True Then TextoFiltro = "TipoConta = 'P'" Else TextoFiltro = "TipoConta = 'R'"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Subfamilia from Familia_financeiro where " & TextoFiltro & " and Subfamilia is not null and ID_PC is null group by Subfamilia", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            .AddItem TBAbrir!SubFamilia
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
1:
    If Permitido = True Then
        If Opt_pagar.Value = True Then TextoFiltro = "Destino = 'P'" Else TextoFiltro = "Destino = 'R'"
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from tbl_familia where " & TextoFiltro & " and CODIGO is not null order by txt_descricao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                Select Case TBFI!Nivel
                    Case 1: TextoFiltro1 = "Left(Codigo,1) = '" & Left(TBFI!CODIGO, 1) & "'"
                    Case 2: TextoFiltro1 = "Left(Codigo,4) = '" & Left(TBFI!CODIGO, 4) & "'"
                    Case 3: TextoFiltro1 = "Left(Codigo,7) = '" & Left(TBFI!CODIGO, 7) & "'"
                    Case 4: TextoFiltro1 = "Left(Codigo,10) = '" & Left(TBFI!CODIGO, 10) & "'"
                    Case 5: TextoFiltro1 = "Left(Codigo,13) = '" & Left(TBFI!CODIGO, 13) & "'"
                    Case 6: TextoFiltro1 = "Left(Codigo,16) = '" & Left(TBFI!CODIGO, 16) & "'"
                    Case 7: TextoFiltro1 = "Left(Codigo,19) = '" & Left(TBFI!CODIGO, 19) & "'"
                    Case 8: TextoFiltro1 = "Left(Codigo,22) = '" & Left(TBFI!CODIGO, 22) & "'"
                End Select
                
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_familia where " & TextoFiltro & " and " & TextoFiltro1 & " order by Nivel", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    TBFIltro.MoveLast
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_familia where " & TextoFiltro & " and int_codfamilia = " & TBFI!int_codfamilia & " and Nivel = " & TBFIltro!Nivel, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
                            .AddItem TBFI!Txt_descricao & " - " & TBFI!CODIGO
                        Else
                            .AddItem TBFI!Txt_descricao
                        End If
                        .ItemData(.NewIndex) = TBFI!int_codfamilia
                    End If
                    TBAbrir.Close
                End If
                TBFIltro.Close
                TBFI.MoveNext
            Loop
        End If
    End If
End With

Exit Sub
tratar_erro:
    If Err.Number = -2147217900 Then
        Permitido = True
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboPara()
On Error GoTo tratar_erro

With Cmb_para
    .Clear
    If Opt_pagar.Value = True Then TextoFiltro = "Destino = 'P'" Else TextoFiltro = "Destino = 'R'"
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_familia where " & TextoFiltro & " and CODIGO is not null order by txt_descricao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Do While TBFI.EOF = False
            Select Case TBFI!Nivel
                Case 1: TextoFiltro1 = "Left(Codigo,1) = '" & Left(TBFI!CODIGO, 1) & "'"
                Case 2: TextoFiltro1 = "Left(Codigo,4) = '" & Left(TBFI!CODIGO, 4) & "'"
                Case 3: TextoFiltro1 = "Left(Codigo,7) = '" & Left(TBFI!CODIGO, 7) & "'"
                Case 4: TextoFiltro1 = "Left(Codigo,10) = '" & Left(TBFI!CODIGO, 10) & "'"
                Case 5: TextoFiltro1 = "Left(Codigo,13) = '" & Left(TBFI!CODIGO, 13) & "'"
                Case 6: TextoFiltro1 = "Left(Codigo,16) = '" & Left(TBFI!CODIGO, 16) & "'"
                Case 7: TextoFiltro1 = "Left(Codigo,19) = '" & Left(TBFI!CODIGO, 19) & "'"
                Case 8: TextoFiltro1 = "Left(Codigo,22) = '" & Left(TBFI!CODIGO, 22) & "'"
            End Select
            
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select * from tbl_familia where " & TextoFiltro & " and " & TextoFiltro1 & " order by Nivel", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                TBFIltro.MoveLast
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_familia where " & TextoFiltro & " and int_codfamilia = " & TBFI!int_codfamilia & " and Nivel = " & TBFIltro!Nivel, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
                        .AddItem TBFI!Txt_descricao & " - " & TBFI!CODIGO
                    Else
                        .AddItem TBFI!Txt_descricao
                    End If
                    .ItemData(.NewIndex) = TBFI!int_codfamilia
                End If
                TBAbrir.Close
            End If
            TBFIltro.Close
            TBFI.MoveNext
        Loop
    End If
End With

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

Private Sub Opt_pagar_Click()
On Error GoTo tratar_erro

ProcCarregaComboDe
ProcCarregaComboPara

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_receber_Click()
On Error GoTo tratar_erro

ProcCarregaComboDe
ProcCarregaComboPara

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcGravar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
