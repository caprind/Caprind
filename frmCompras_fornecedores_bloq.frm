VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_fornecedores_bloq 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Compras - Fornecedores - Status"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   5550
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_fornecedores_bloq.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2895
      Left            =   55
      TabIndex        =   3
      Top             =   990
      Width           =   6330
      Begin VB.TextBox txtStatus 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Status."
         Top             =   180
         Width           =   4845
      End
      Begin VB.TextBox txtObservacoes 
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
         Height          =   1875
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Observações."
         Top             =   900
         Width           =   4845
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   540
         Width           =   4845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observações :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   900
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Width           =   1035
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   6330
      _ExtentX        =   11165
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   44
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Bloquear"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Bloquear (F6)"
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
      ButtonLeft2     =   48
      ButtonTop2      =   2
      ButtonWidth2    =   58
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Desbloquear"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Desbloquear (F7)"
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
      ButtonLeft3     =   108
      ButtonTop3      =   2
      ButtonWidth3    =   79
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
      ButtonLeft4     =   189
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   193
      ButtonTop5      =   2
      ButtonWidth5    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   236
      ButtonTop6      =   2
      ButtonWidth6    =   30
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   268
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
   End
End
Attribute VB_Name = "frmCompras_fornecedores_bloq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextoFiltro_Instituicao As String 'OK

Private Sub ProcBloquear()
On Error GoTo tratar_erro

With txtObservacoes
    .Locked = False
    .TabStop = True
    .Text = ""
    .SetFocus
End With
If Financeiro_Instituicao = True Then txtStatus.Text = "Bloqueada" Else txtStatus.Text = "Bloqueado"
txtResponsavel.Text = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesbloquear()
On Error GoTo tratar_erro

With txtObservacoes
    .Locked = False
    .TabStop = True
    .Text = ""
    .SetFocus
End With
If Financeiro_Instituicao = True Then txtStatus.Text = "Liberada" Else txtStatus.Text = "Liberado"
txtResponsavel = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If txtStatus = "" Then
    USMsgBox ("Informe se o registro está bloqueado ou liberado antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Financeiro_Instituicao = True Then
    With frm_Instituicoes
        For InitFor = 1 To .lst_Instituicoes.ListItems.Count
            If .lst_Instituicoes.ListItems.Item(InitFor).Checked = True Then
                TextoFiltro_Instituicao = "Bloqueado, Txt_descricao, ID from tbl_Instituicoes where ID = " & .lst_Instituicoes.ListItems.Item(InitFor)
                ProcEnviaDados
            End If
        Next InitFor
    
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Instituicoes where id = " & IIf(.txtCodBanco = "", 0, .txtCodBanco), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then .ProcCarregaDados
        TBAbrir.Close
        .ProcCarregaLista
    End With
Else
    With frmCompras_fornecedores
        For InitFor = 1 To .Lista.ListItems.Count
            If .Lista.ListItems.Item(InitFor).Checked = True Then
                TextoFiltro_Instituicao = "Status, Nome_Razao, IDcliente from Compras_fornecedores where IDcliente = " & .Lista.ListItems.Item(InitFor)
                ProcEnviaDados
            End If
        Next InitFor

        Set TBFornecedor = CreateObject("adodb.recordset")
        TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
        If TBFornecedor.EOF = False Then .ProcPuxaDados
        TBFornecedor.Close
        .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    End With
End If
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF6: ProcBloquear
    Case vbKeyF7: ProcDesbloquear
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

ProcCarregaToolBar1 Me, 6330, 6, True

contador = 0
If Financeiro_Instituicao = True Then
    Caption = "Administrativo - Financeiro - Instituições - Status"
    With frm_Instituicoes
        For InitFor = 1 To .lst_Instituicoes.ListItems.Count
            If .lst_Instituicoes.ListItems.Item(InitFor).Checked = True Then
                contador = contador + 1
            End If
        Next InitFor
        If contador = 1 Then TextoFiltro_Instituicao = "Bloqueado from tbl_Instituicoes where ID = " & IIf(.txtCodBanco = "", 0, .txtCodBanco)
    End With
Else
    Caption = "Administrativo - Compras - Fornecedores - Status"
    With frmCompras_fornecedores
        For InitFor = 1 To .Lista.ListItems.Count
            If .Lista.ListItems.Item(InitFor).Checked = True Then
                contador = contador + 1
            End If
        Next InitFor
        If contador = 1 Then TextoFiltro_Instituicao = "Status from Compras_fornecedores where IDCliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente)
    End With
End If
If contador = 1 Then
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select resp_Status, obs_Status, " & TextoFiltro_Instituicao, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        If Financeiro_Instituicao = True Then
            txtStatus.Text = IIf(TBFornecedor!Bloqueado = False, "Liberada", "Bloqueada")
        Else
            txtStatus.Text = IIf(IsNull(TBFornecedor!status), "", TBFornecedor!status)
        End If
        txtObservacoes.Text = IIf(IsNull(TBFornecedor!obs_Status), "", TBFornecedor!obs_Status)
        txtResponsavel.Text = IIf(IsNull(TBFornecedor!resp_Status), "", TBFornecedor!resp_Status)
    End If
    TBFornecedor.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcBloquear
    Case 3: ProcDesbloquear
    'Case 5: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcEnviaDados()
On Error GoTo tratar_erro

Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select obs_Status, resp_Status, " & TextoFiltro_Instituicao, Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    If txtStatus.Text = "Liberada" Or txtStatus.Text = "Liberado" Then
        If Financeiro_Instituicao = True Then
            TBFornecedor!Bloqueado = False
            Evento = "Liberar instituição"
        Else
            TBFornecedor!status = "Liberado"
            Evento = "Liberar fornecedor"
        End If
    Else
        If Financeiro_Instituicao = True Then
            TBFornecedor!Bloqueado = True
            Evento = "Bloquear instituição"
        Else
            TBFornecedor!status = "Bloqueado"
            Evento = "Bloquear fornecedor"
        End If
    End If
    TBFornecedor!obs_Status = IIf(txtObservacoes = "", Null, txtObservacoes)
    TBFornecedor!resp_Status = txtResponsavel.Text
    TBFornecedor.Update
    
    '==================================
    If Financeiro_Instituicao = True Then
        Modulo = "Financeiro/Instituições"
        ID_documento = TBFornecedor!ID
        Documento = "Fornecedor: " & TBFornecedor!Txt_descricao
    Else
        Modulo = "Compras/Fornecedores"
        ID_documento = TBFornecedor!IDCliente
        Documento = "Fornecedor: " & TBFornecedor!Nome_Razao
    End If
    Documento1 = ""
    ProcGravaEvento
    '==================================
End If
TBFornecedor.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
