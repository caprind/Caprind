VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmRequisicao_materiais_Autorizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RM - Selecionar autorizado"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4065
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmRequisicao_materiais_Autorizar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2490
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmRequisicao_materiais_Autorizar.frx":0442
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   55
      TabIndex        =   2
      Top             =   990
      Width           =   3975
      Begin VB.TextBox txtsenha 
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
         IMEMode         =   3  'DISABLE
         Left            =   900
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Senha."
         Top             =   540
         Width           =   2865
      End
      Begin VB.TextBox cmbaut 
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
         Left            =   900
         TabIndex        =   0
         ToolTipText     =   "Usuário."
         Top             =   180
         Width           =   2865
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Senha :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuário :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Nº:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -11580
         TabIndex        =   3
         Top             =   4200
         Width           =   270
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
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
      ButtonCaption1  =   "Autorizar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Autorizar (F3)"
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
      ButtonWidth1    =   52
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
      ButtonLeft2     =   56
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
      ButtonLeft3     =   60
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
      ButtonLeft4     =   103
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
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
End
Attribute VB_Name = "frmRequisicao_materiais_Autorizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub procAutorizar()
On Error GoTo tratar_erro

Acao = "autorizar"
If cmbaut.Text = "" Then
    NomeCampo = "o usuário"
    ProcVerificaAcao
    cmbaut.SetFocus
    Exit Sub
Else
    cmbaut.Text = UCase(cmbaut.Text)
End If
If txtSenha.Text = "" Then
    NomeCampo = "a senha"
    ProcVerificaAcao
    txtSenha.SetFocus
    Exit Sub
End If

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select IDUsuario, Bloqueado from usuarios where usuario = '" & cmbaut & "' and Senha = '" & txtSenha & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    If TBUsuarios!Bloqueado = True Then
        USMsgBox ("Não é permitido efetuar a autorização, pois o usuário " & cmbaut.Text & " está bloqueado."), vbExclamation, "CAPRIND v5.0"
        TBUsuarios.Close
        Exit Sub
    End If
    
    'Verifica se o usuário tem acesso liberado do CC
    INNERJOINTEXTO = "(Requisicao_materiais_lista RML INNER JOIN Usuarios_Setor CC ON CC.ID = RML.ID_CC) INNER JOIN Usuarios_Setor_Responsavel CCR ON CCR.ID_CC = CC.Id and CCR.Responsavel_CC = '" & cmbaut & "'"
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Acessos where IDusuario = " & TBUsuarios!IDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then INNERJOINTEXTO = "Requisicao_materiais_lista RML"
    TBAbrir.Close
Else
    USMsgBox ("Verifique se o usuário e a senha estão corretos."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
TBUsuarios.Close

Permitido = False
With frmRequisicao_materiais
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select RML.Data_autorizacao, RML.Autorizado, RML.Desenho from " & INNERJOINTEXTO & " where RML.IdLista = " & .Lista.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras.EOF = False Then
                Permitido = True
                If IsNull(TBCompras!Data_autorizacao) = True Then
                    TBCompras!Autorizado = IIf(cmbaut = "", Null, cmbaut)
                    TBCompras!Data_autorizacao = Date
                    Evento = "Autorizar"
                Else
                    TBCompras!Autorizado = Null
                    TBCompras!Data_autorizacao = Null
                    Evento = "Cancelar autorização"
                End If
                TBCompras.Update
                
                '==================================
                Modulo = "Estoque/Requisição de materiais"
                ID_documento = .Lista.ListItems.Item(InitFor)
                Documento = "Nº requisição: " & .txtrequisicao
                Documento1 = "Cód. interno: " & TBCompras!Desenho
                ProcGravaEvento
                '==================================
            End If
        End If
Proximo:
        'TBCompras.Close
    Next InitFor
    If Permitido = True Then
        USMsgBox ("Autorização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Requisicao_materiais_lista where IdLista = " & .Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            .procCarregaDados_Itens
        End If
        TBAbrir.Close
    Else
        USMsgBox ("Autorização não efetuada, pois o usuário não é responsável pelo(s) centro(s) de custo."), vbExclamation, "CAPRIND v5.0"
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbaut_LostFocus()
On Error GoTo tratar_erro

If cmbaut <> "" Then cmbaut.Text = UCase(cmbaut.Text)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF3: procAutorizar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 3975, 5, True
Frame1.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Frame1.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procAutorizar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

