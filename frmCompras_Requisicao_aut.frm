VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_Requisicao_aut 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cotação - Selecionar autorizado"
   ClientHeight    =   1995
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
   Icon            =   "frmCompras_Requisicao_aut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2490
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_Requisicao_aut.frx":0442
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   975
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
Attribute VB_Name = "frmCompras_Requisicao_aut"
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
End If
If txtSenha.Text = "" Then
    NomeCampo = "a senha"
    ProcVerificaAcao
    txtSenha.SetFocus
    Exit Sub
End If

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * from usuarios where usuario = '" & cmbaut.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    UsuarioTexto = TBUsuarios!Usuario
    
    If Compras_Cotacao = True Then
        TextoFiltro = "acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = 'Compras/Cotação/Liberar cotação'"
        Mensagem_acesso = "liberar a cotação"
        Mensagem_acesso1 = "liberação"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select ID_CC from Requisicao_materiais_lista where IdLista = " & frmRequisicao_materiais.TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TextoFiltro = "Usuarios_Setor_Responsavel where ID_CC = " & TBAbrir!ID_CC & " and Responsavel_CC = '" & cmbaut & "'"
            Mensagem_acesso = "autorizar a requisição deste produto neste centro de custo"
            Mensagem_acesso1 = "liberação"
        End If
        
        'Verifica se o usuário tem acesso liberado do CC
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Acessos where IDusuario = " & TBUsuarios!IDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then GoTo Pular
        TBAbrir.Close
    End If
    Set TBAcessos = CreateObject("adodb.recordset")
    TBAcessos.Open "select * from " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBAcessos.EOF = True Then
        USMsgBox ("Atenção usuário " & cmbaut & ", você não tem autorização para " & Mensagem_acesso & "."), vbExclamation, "CAPRIND v5.0"
        TBAcessos.Close
        Exit Sub
    End If
    TBAcessos.Close
    If TBUsuarios!Bloqueado = True Then
        USMsgBox ("Não é permitido efetuar a " & Mensagem_acesso1 & ", pois o usuário " & cmbaut.Text & " está bloqueado."), vbExclamation, "CAPRIND v5.0"
        TBUsuarios.Close
        Exit Sub
    End If
Else
    USMsgBox ("Verifique se o usuário e a senha estão corretos."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
TBUsuarios.Close

Pular:
    If Compras_Cotacao = True Then
        With frmcompras_reqcot
            Cont = .txtidcot
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from compras_cotacao where id_cotacao = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
            TBCompras!Autorizado = UsuarioTexto
            If .txtStatus <> "LIBERADA" Then
                TBCompras!statuscotacao = "LIBERADA"
                TBCompras!dataliberada = Format(Date, "dd/mm/yyyy")
                USMsgBox ("Cotação liberada com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "Compras/Cotação"
                Evento = "Liberar"
                ID_documento = Cont
                Documento = "Nº cotação: " & .txtidcotacao
                Documento1 = ""
                ProcGravaEvento
                '==================================
            Else
                TBCompras!statuscotacao = "COTANDO"
                TBCompras!dataliberada = Null
                USMsgBox ("Liberação cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
            End If
            .txtStatus = TBCompras!statuscotacao
            TBCompras.Update
            '==================================
            Modulo = "Compras/Cotação"
            Evento = "Cancelar liberação"
            ID_documento = Cont
            Documento = "Nº cotação: " & .txtidcotacao
            Documento1 = ""
            ProcGravaEvento
            '==================================
            .lista_cot.ListItems.Clear
            .ProcCarregaListaCotacao (IIf(ReturnNumbersOnly(Left(.lblPaginas(2).Caption, Len(.lblPaginas(2).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(2).Caption, Len(.lblPaginas(2).Caption) - 5))))
        End With
    Else
        With frmRequisicao_materiais
            If .txtAutorizado.Text = "" Then
                .txtAutorizado.Text = UsuarioTexto
                .txtData_Autorizacao.Text = Format(Date, "dd/mm/yy")
                USMsgBox ("Autorização efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
                Evento = "Autorizar"
            Else
                .txtAutorizado.Text = ""
                .txtData_Autorizacao.Text = ""
                USMsgBox ("Autorização cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
                Evento = "Cancelar autorização"
            End If
            Set TBCompras = CreateObject("adodb.recordset")
            TBCompras.Open "Select * from Requisicao_materiais_lista where IdLista = " & frmRequisicao_materiais.TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
            TBCompras!Autorizado = IIf(.txtAutorizado.Text = "", Null, .txtAutorizado)
            TBCompras!Data_autorizacao = IIf(.txtData_Autorizacao.Text = "", Null, Format(Date, "dd/mm/yy"))
            TBCompras.Update
            TBCompras.Close
            '==================================
            Modulo = "Estoque/Requisição de materiais"
            ID_documento = .TXTIDLista
            Documento = "Nº requisição: " & .txtrequisicao
            Documento1 = "Cód. interno: " & .txtdesenho
            ProcGravaEvento
            '==================================
        End With
    End If
    Unload Me

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

With cmbaut
    Set TBAbrir = CreateObject("adodb.recordset")
    If Compras_Cotacao = True Then
        Cont = frmcompras_reqcot.txtidcot
        TBAbrir.Open "Select * from Compras_Cotacao where id_cotacao = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If TBAbrir!dataliberada <> "" Or IsNull(TBAbrir!dataliberada) = False Then
                .Text = TBAbrir!Autorizado
                .Locked = True
                .TabStop = False
            End If
        End If
        frmCompras_Requisicao_aut.Caption = "Cotação - Selecionar autorizado"
    Else
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Requisicao_materiais_lista where IdLista = " & frmRequisicao_materiais.TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If TBAbrir!Data_autorizacao <> "" Or IsNull(TBAbrir!Data_autorizacao) = False Then
                .Text = TBAbrir!Autorizado
                .Locked = True
                .TabStop = False
            End If
        End If
        frmCompras_Requisicao_aut.Caption = "Requisição de materiais - Selecionar autorizado"
    End If
    TBAbrir.Close
End With
Frame1.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcBloqueiaCampos()
On Error GoTo tratar_erro

With frmCompras_Requisicao
    .txtobssolicitacao.Locked = True
    .txtobssolicitacao.TabStop = False
End With

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

Sub Procliberacampos()
On Error GoTo tratar_erro

With frmCompras_Requisicao
    .txtobssolicitacao.Locked = False
    .txtobssolicitacao.TabStop = True
End With

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

