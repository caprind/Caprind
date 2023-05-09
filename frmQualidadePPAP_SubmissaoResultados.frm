VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmQualidadePPAP_SubmissaoResultados 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - PPAP - PSW - Submissão - Resultados"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7830
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   6000
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmQualidadePPAP_SubmissaoResultados.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3285
      Left            =   55
      TabIndex        =   7
      Top             =   990
      Width           =   7725
      Begin VB.CheckBox chk4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados estatísticos do processo"
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
         TabIndex        =   3
         Top             =   1530
         Width           =   2535
      End
      Begin VB.CheckBox chk3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Critérios de aparência"
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
         TabIndex        =   2
         Top             =   1100
         Width           =   1875
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medição dimensionais"
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
         Top             =   240
         Width           =   1845
      End
      Begin VB.CheckBox chk2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ensaios de materiais e funcionais"
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
         TabIndex        =   1
         Top             =   670
         Width           =   2655
      End
      Begin VB.CheckBox chkSim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sim"
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
         Left            =   6180
         TabIndex        =   4
         Top             =   2010
         Width           =   555
      End
      Begin VB.CheckBox chkNao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não"
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
         TabIndex        =   5
         Top             =   2010
         Width           =   585
      End
      Begin VB.TextBox txtOBS 
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
         Height          =   495
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Molde/Cavidade/Processos de produção."
         Top             =   2640
         Width           =   7305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Molde/Cavidade/Processos de produção"
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
         Left            =   2385
         TabIndex        =   9
         Top             =   2430
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estes resultados atendem a todos os requisitos do desenho(s) e especificações?"
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
         TabIndex        =   8
         Top             =   2010
         Width           =   5760
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   10
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
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
Attribute VB_Name = "frmQualidadePPAP_SubmissaoResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkNao_Click()
On Error GoTo tratar_erro

If chkNao.Value = 1 Then
    chkSim.Value = 0
    txtobs.Locked = False
    txtobs.TabStop = True
Else
    txtobs.Locked = True
    txtobs.TabStop = False
    txtobs = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkSim_Click()
On Error GoTo tratar_erro

If chkSim.Value = 1 Then
    chkNao.Value = 0
    txtobs.Locked = True
    txtobs.TabStop = False
    txtobs = ""
End If

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

ProcCarregaToolBar1 Me, 7725, 5, True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT * from QualidadePPAP where IDPPAP = " & frmQualidadePPAP.txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Chk1_Resultados = False Then Chk1.Value = 0 Else Chk1.Value = 1
    If TBAbrir!Chk2_Resultados = False Then Chk2.Value = 0 Else Chk2.Value = 1
    If TBAbrir!Chk3_Resultados = False Then Chk3.Value = 0 Else Chk3.Value = 1
    If TBAbrir!Chk4_Resultados = False Then Chk4.Value = 0 Else Chk4.Value = 1
    If TBAbrir!Sim_Resultados = False Then chkSim.Value = 0 Else chkSim.Value = 1
    If TBAbrir!Nao_Resultados = False Then
        chkNao.Value = 0
    Else
        chkNao.Value = 1
        txtobs = IIf(IsNull(TBAbrir!Obs_Resultados), "", TBAbrir!Obs_Resultados)
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerifValidacaoRegistro("alterar", frmQualidadePPAP.txtData_Validacao, "mesmo", "o PPAP", True) = False Then Exit Sub
With frmQualidadePPAP
    If .txtStatus.Visible = True Then
        USMsgBox ("Não é permitida a alteração do PPAP revisado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "SELECT * from QualidadePPAP where IDPPAP = " & .txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        If Chk1.Value = False Then TBGravar!Chk1_Resultados = False Else TBGravar!Chk1_Resultados = True
        If Chk2.Value = False Then TBGravar!Chk2_Resultados = False Else TBGravar!Chk2_Resultados = True
        If Chk3.Value = False Then TBGravar!Chk3_Resultados = False Else TBGravar!Chk3_Resultados = True
        If Chk4.Value = False Then TBGravar!Chk4_Resultados = False Else TBGravar!Chk4_Resultados = True
        If chkSim.Value = False Then TBGravar!Sim_Resultados = False Else TBGravar!Sim_Resultados = True
        If chkNao.Value = False Then
            TBGravar!Nao_Resultados = False
            TBGravar!Obs_Resultados = ""
        Else
            TBGravar!Nao_Resultados = True
            TBGravar!Obs_Resultados = txtobs
        End If
        TBGravar.Update
        USMsgBox ("Resultados cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/PPAP/PSW/Submissão/Resultados"
        Evento = "Cadastrar resultados"
        ID_documento = .txtIDPPAP
        Documento = "Número PPAP: " & .txtPPAP.Text & " - Cód. interno: " & .txtCodinterno
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
    TBGravar.Close
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
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
