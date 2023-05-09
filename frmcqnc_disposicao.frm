VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmcqnc_disposicao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - Não conformidade - Disposição"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15315
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2325
      Left            =   55
      TabIndex        =   18
      Top             =   960
      Width           =   15195
      Begin VB.TextBox TxtdescricaoNC 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Descrição da não conformidade."
         Top             =   390
         Width           =   5715
      End
      Begin VB.CheckBox chkAnalizada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Analisada"
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
         Left            =   14010
         TabIndex        =   16
         Top             =   1890
         Width           =   1005
      End
      Begin VB.TextBox txtobscq 
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
         Height          =   465
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         ToolTipText     =   "Observação do parecer do controle de qualidade."
         Top             =   1740
         Width           =   13725
      End
      Begin VB.CommandButton cmdPFP 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8550
         Picture         =   "frmcqnc_disposicao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Localizar dimensões."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtParecerF 
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
         Left            =   2970
         TabIndex        =   2
         ToolTipText     =   "Dimensão."
         Top             =   390
         Width           =   5565
      End
      Begin VB.ComboBox Cmb_origem 
         Appearance      =   0  'Flat
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Origem."
         Top             =   390
         Width           =   2325
      End
      Begin VB.CommandButton Cmd_cadastrar_origem 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2520
         Picture         =   "frmcqnc_disposicao.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cadastrar origem."
         Top             =   390
         Width           =   315
      End
      Begin VB.CommandButton Cmd_cadastrar_causa 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   14700
         Picture         =   "frmcqnc_disposicao.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cadastrar causa."
         Top             =   390
         Width           =   315
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disposição*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   180
         TabIndex        =   19
         Top             =   810
         Width           =   14865
         Begin VB.OptionButton Opt_nada_consta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nada consta"
            DisabledPicture =   "frmcqnc_disposicao.frx":0306
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
            Left            =   13530
            TabIndex        =   14
            Top             =   300
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Opt_aprovado_desvio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprovado com desvio"
            DisabledPicture =   "frmcqnc_disposicao.frx":24A248
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
            Left            =   1338
            TabIndex        =   7
            Top             =   300
            Width           =   1845
         End
         Begin VB.OptionButton Opt_outros 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Outros"
            DisabledPicture =   "frmcqnc_disposicao.frx":49418A
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
            Left            =   12576
            TabIndex        =   13
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton Opt_reaproveitar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reaproveitar para outro produto"
            DisabledPicture =   "frmcqnc_disposicao.frx":6DE0CC
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
            Left            =   6870
            TabIndex        =   11
            Top             =   300
            Width           =   2685
         End
         Begin VB.OptionButton Opt_aprovado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aprovado"
            DisabledPicture =   "frmcqnc_disposicao.frx":92800E
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
            TabIndex        =   6
            Top             =   300
            Width           =   1005
         End
         Begin VB.OptionButton Opt_selecionar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Selecionar"
            DisabledPicture =   "frmcqnc_disposicao.frx":B71F50
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
            Height          =   210
            Left            =   5682
            TabIndex        =   10
            Top             =   300
            Width           =   1035
         End
         Begin VB.OptionButton Opt_retrabalhar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Retrabalhar"
            DisabledPicture =   "frmcqnc_disposicao.frx":DBBE92
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
            Height          =   210
            Left            =   4374
            TabIndex        =   9
            Top             =   300
            Width           =   1155
         End
         Begin VB.OptionButton Opt_rejeitar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rejeitar"
            DisabledPicture =   "frmcqnc_disposicao.frx":1005DD4
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
            Height          =   210
            Left            =   3336
            TabIndex        =   8
            Top             =   300
            Width           =   885
         End
         Begin VB.OptionButton optDevolver 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Devolver para fornecedor/cliente"
            DisabledPicture =   "frmcqnc_disposicao.frx":124FD16
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
            Left            =   9708
            TabIndex        =   12
            Top             =   300
            Width           =   2715
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensão"
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
         Left            =   5407
         TabIndex        =   23
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição da não conformidade"
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
         Left            =   10702
         TabIndex        =   22
         Top             =   180
         Width           =   2250
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
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
         Left            =   6570
         TabIndex        =   21
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Origem"
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
         Left            =   1087
         TabIndex        =   20
         Top             =   180
         Width           =   510
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7320
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmcqnc_disposicao.frx":1499C58
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   17
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
         Name            =   "Tahoma"
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
      ButtonUseMaskColor2=   0   'False
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
   End
End
Attribute VB_Name = "frmcqnc_disposicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_cadastrar_causa_Click()
On Error GoTo tratar_erro

Sit_REG = 1
frmcqnc_causa.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cadastrar_origem_Click()
On Error GoTo tratar_erro

Sit_REG = 1
frmcqnc_origem.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPFP_Click()
On Error GoTo tratar_erro

Sit_REG = 1
frmcqnc_dimensoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
    Case vbKeyF3: ProcSalvar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
frmcqnc.ProcCarregaComboOrigem Cmb_origem

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

Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"

If Opt_aprovado.Value = False And Opt_aprovado_desvio.Value = False And Opt_rejeitar.Value = False And Opt_retrabalhar.Value = False And Opt_selecionar.Value = False And Opt_reaproveitar.Value = False And optDevolver.Value = False And Opt_outros.Value = False And Opt_nada_consta.Value = False Then
    NomeCampo = "a disposição"
    ProcVerificaAcao
    Exit Sub
End If
If Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True Or Opt_rejeitar.Value = True Or Opt_retrabalhar.Value = True Or Opt_selecionar.Value = True Or Opt_reaproveitar.Value = True Or optDevolver.Value = True Or Opt_outros.Value = True Then
    If USMsgBox("Essa(s) não conformidade(s) já foi(ram) analisada(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then chkAnalizada.Value = 1 Else chkAnalizada.Value = 0
End If

CamposLoop = ""
With frmcqnc.ListaFases
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select CQNCF.*, P.Desenho from CQ_NC_FABRICA CQNCF INNER JOIN Producao P ON CQNCF.Ordem = P.Ordem where CQNCF.Codigo = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                If txtParecerF.Text <> "" Then TBGravar!PARECERFAB = txtParecerF.Text
                If Cmb_origem <> "" Then TBGravar!ID_origem = Cmb_origem.ItemData(Cmb_origem.ListIndex)
                If txtParecerF <> "" Then TBGravar!obsFab = TxtdescricaoNC
                
                ProcGravarNCOSMaq

                If Opt_nada_consta.Value = True Then
                    TBGravar!ParecerCQ = "Nada consta"
                    chkAnalizada.Value = 0
                End If
                If Opt_aprovado.Value = True Then TBGravar!ParecerCQ = "Aprovado"
                If Opt_aprovado_desvio.Value = True Then TBGravar!ParecerCQ = "Aprovado c/ desvio"
                If Opt_rejeitar.Value = True Then TBGravar!ParecerCQ = "Rejeitar"
                If Opt_reaproveitar.Value = True Then TBGravar!ParecerCQ = "Reaproveitar"
                If optDevolver.Value = True Then TBGravar!ParecerCQ = "Devolver"
                If Opt_outros.Value = True Then TBGravar!ParecerCQ = "Outros"
                If Opt_retrabalhar.Value = True Then TBGravar!ParecerCQ = "Retrabalhar"
                
                If chkAnalizada.Value = 1 Then TBGravar!Analizada = True Else TBGravar!Analizada = False
                
                TBGravar!obsCQ = txtobscq.Text
                
                TBGravar.Update

                If IsNull(TBGravar!Ordem) = False Then frmcqnc.ProcGravarNCOrdem TBGravar!Ordem
                
                '==================================
                Documento = "Data: " & Format(TBGravar!Data, "dd/mm/yy") & " - Hora: " & Format(TBGravar!Hora, "hh:mm:ss") & " - Ordem: " & TBGravar!Ordem & " - OS: " & TBGravar!OS & " - Cód. interno: " & TBGravar!Desenho & " - Operador: " & TBGravar!Operador
                If CamposLoop = "" Then
                    CamposLoop = "('Qualidade/Não conformidade','" & pubUsuario & "','Alterar disposição','" & Date & "','" & Time & "','" & Documento & "'," & .ListItems(InitFor) & ")"
                Else
                    CamposLoop = CamposLoop & ", ('Qualidade/Não conformidade','" & pubUsuario & "','Alterar disposição','" & Date & "','" & Time & "','" & Documento & "'," & .ListItems(InitFor) & ")"
                End If
                '==================================
            End If
            TBGravar.Close
        End If
    Next InitFor
End With

'==================================
Conexao.Execute "INSERT INTO Mascara (Modulo, Usuario, Operacao, Data, Hora, Documento, ID_documento) VALUES " & CamposLoop
'==================================

USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
Unload Me

With frmcqnc
    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    If CodigoLista <> 0 And .ListaFases.ListItems.Count <> 0 Then
        .ListaFases.SelectedItem = .ListaFases.ListItems(CodigoLista)
        .ListaFases.SetFocus
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcGravarNCOSMaq()
On Error GoTo tratar_erro

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Ordemservico_maq_utilizadas where OS = " & TBGravar!OS & " and Maquina = '" & TBGravar!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    QTNC = 0
    Set TBProducaoFases = CreateObject("adodb.recordset")
    TBProducaoFases.Open "Select Sum(TTNC) as QTNC from CQ_NC_FABRICA where Codigo <> " & TBGravar!CODIGO & " and OS = " & TBGravar!OS & " and Maquina = '" & TBGravar!maquina & "' and (PARECERCQ = 'Rejeitar' or PARECERCQ = 'Retrabalhar' or PARECERCQ = 'Selecionar' or PARECERCQ = 'Outros' or PARECERCQ = 'Nada consta')", Conexao, adOpenKeyset, adLockOptimistic
    If TBProducaoFases.EOF = False Then
        QTNC = IIf(IsNull(TBProducaoFases!QTNC), 0, TBProducaoFases!QTNC)
    End If
    Set TBAfericao = CreateObject("adodb.recordset")
    TBAfericao.Open "Select * FROM producao INNER JOIN Ordemservico ON Ordemservico.Ordem = producao.Ordem where Ordemservico.IDProducao = " & TBGravar!OS & " and Producao.Ap_backup = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAfericao.EOF = False Then
        NomeTabelaAp = "ProducaoFases_Backup"
    Else
        NomeTabelaAp = "ProducaoFases"
    End If
    
    'Atualiza dados no apontamento
    If IsNull(TBGravar!IDProducao) = False And TBGravar!IDProducao <> "0" Then
        QTNCAP = 0
        Set TBProducaoFases = CreateObject("adodb.recordset")
        TBProducaoFases.Open "Select Sum(TTNC) as QTNCAP from CQ_NC_FABRICA where Codigo <> " & TBGravar!CODIGO & " and IDproducao = " & TBGravar!IDProducao & " and (PARECERCQ = 'Rejeitar' or PARECERCQ = 'Retrabalhar' or PARECERCQ = 'Selecionar' or PARECERCQ = 'Outros' or PARECERCQ = 'Nada consta')", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducaoFases.EOF = False Then
            QTNCAP = IIf(IsNull(TBProducaoFases!QTNCAP), 0, TBProducaoFases!QTNCAP)
        End If
    End If
    
    If IsNull(TBGravar!ParecerCQ) = True Or TBGravar!ParecerCQ = "" Or TBGravar!ParecerCQ = "Nada consta" Then
        If Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True Or Opt_reaproveitar.Value = True Then
            TBproducao!QTOK = TBproducao!QTOK + IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC)
            TBproducao!QTNC = IIf(TBproducao!QTNC - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC) < 0, 0, TBproducao!QTNC - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC))

            'Atualiza dados no apontamento
            If IsNull(TBGravar!IDProducao) = False And TBGravar!IDProducao <> "0" Then
                Set TBProducaoFases = CreateObject("adodb.recordset")
                TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                If TBProducaoFases.EOF = False Then
                    TBProducaoFases!Reprovada = IIf(TBProducaoFases!Reprovada - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC) < 0, 0, QTNCAP - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC))
                    TBProducaoFases!quantidade = TBProducaoFases!quantidade + IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC)
                    TBProducaoFases.Update
                    
                    frmcqnc.ProcAtualizaMovEstoque
                End If
                TBProducaoFases.Close
            End If
        End If
    Else
        If TBGravar!ParecerCQ <> "Rejeitar" And TBGravar!ParecerCQ <> "Retrabalhar" And TBGravar!ParecerCQ <> "Selecionar" And TBGravar!ParecerCQ <> "Outros" And (Opt_rejeitar.Value = True Or Opt_retrabalhar.Value = True Or Opt_selecionar.Value = True Or Opt_outros.Value = True Or Opt_nada_consta.Value = True) Then
            TBproducao!QTNC = QTNC + IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC)
            TBproducao!QTOK = IIf(TBproducao!QTOK - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC) < 0, 0, TBproducao!QTOK - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC))

            'Atualiza dados no apontamento
            If TBGravar!IDProducao <> "" And TBGravar!IDProducao <> "0" Then
                Set TBProducaoFases = CreateObject("adodb.recordset")
                TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                If TBProducaoFases.EOF = False Then
                    TBProducaoFases!Reprovada = QTNCAP + IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC)
                    TBProducaoFases!quantidade = IIf(TBProducaoFases!quantidade - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC) < 0, 0, TBProducaoFases!quantidade - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC))
                    TBProducaoFases.Update
                    
                    frmcqnc.ProcAtualizaMovEstoque
                End If
                TBProducaoFases.Close
            End If
        End If
        If TBGravar!ParecerCQ <> "Aprovado" And TBGravar!ParecerCQ <> "Aprovado c/ desvio" And TBGravar!ParecerCQ <> "Reaproveitar" And (Opt_aprovado.Value = True Or Opt_aprovado_desvio.Value = True) Then
            TBproducao!QTNC = IIf(QTNC - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC) < 0, 0, QTNC - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC))
            TBproducao!QTOK = TBproducao!QTOK + IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC)

            'Atualiza dados no apontamento
            If IsNull(TBGravar!IDProducao) = False And TBGravar!IDProducao <> "0" Then
                Set TBProducaoFases = CreateObject("adodb.recordset")
                TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & TBGravar!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                If TBProducaoFases.EOF = False Then
                    TBProducaoFases!Reprovada = IIf(QTNCAP - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC) < 0, 0, QTNCAP - IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC))
                    TBProducaoFases!quantidade = TBProducaoFases!quantidade + IIf(IsNull(TBGravar!TTNC), 0, TBGravar!TTNC)
                    TBProducaoFases.Update
                    
                    frmcqnc.ProcAtualizaMovEstoque
                End If
                TBProducaoFases.Close
            End If
        End If
    End If

    TBproducao!Totalprod = TBproducao!QTOK + TBproducao!QTNC
    TBproducao.Update
End If
TBproducao.Close

frmcqnc.ProcGravarNCOS TBGravar!OS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

