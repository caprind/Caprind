VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmQualidadePPAP_SubmissaoNivel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - PPAP - PSW - Submissão - Nível"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10320
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7050
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmQualidadePPAP_SubmissaoNivel.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   55
      TabIndex        =   5
      Top             =   1020
      Width           =   10185
      Begin VB.OptionButton opt5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nível 5 - Certificado com amostras do produto e dados de suporte verificados na localidade de manufaturada do fornecedor"
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
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   1605
         Width           =   9255
      End
      Begin VB.OptionButton opt4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nível 4 - Certificado e outros requisitos conforme definido pelo cliente"
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
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   1248
         Width           =   5325
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nível 3 - Certificado com amostras do produto e todos os dados de suporte submetido ao cliente"
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
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   892
         Width           =   7245
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nível 2 - Certificado com amostras do produto e dados limitados de suporte submetido ao cliente"
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
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   536
         Width           =   7245
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nível 1 - Certificado apenas (e para itens designados de aparência, um Relatório de Aprovação de Aparência) submetido ao cliente"
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
         Height          =   240
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   9765
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   6
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
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
Attribute VB_Name = "frmQualidadePPAP_SubmissaoNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

ProcCarregaToolBar1 Me, 10185, 5, True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT * from QualidadePPAP where IDPPAP = " & frmQualidadePPAP.txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!opt1_Nivel = False Then opt1 = False Else opt1 = True
    If TBAbrir!opt2_Nivel = False Then opt2 = False Else opt2 = True
    If TBAbrir!opt3_Nivel = False Then opt3 = False Else opt3 = True
    If TBAbrir!opt4_Nivel = False Then opt4 = False Else opt4 = True
    If TBAbrir!opt5_Nivel = False Then opt5 = False Else opt5 = True
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
If opt1.Value = False And opt2.Value = False And opt3.Value = False And opt4.Value = False And opt5.Value = False Then
    USMsgBox "Informe uma das opções antes de salvar.", vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmQualidadePPAP
    If .txtStatus.Visible = True Then
        USMsgBox ("Não é permitida a alteração do PPAP revisado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "SELECT * from QualidadePPAP where IDPPAP = " & .txtIDPPAP, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        If opt1.Value = False Then TBGravar!opt1_Nivel = False Else TBGravar!opt1_Nivel = True
        If opt2.Value = False Then TBGravar!opt2_Nivel = False Else TBGravar!opt2_Nivel = True
        If opt3.Value = False Then TBGravar!opt3_Nivel = False Else TBGravar!opt3_Nivel = True
        If opt4.Value = False Then TBGravar!opt4_Nivel = False Else TBGravar!opt4_Nivel = True
        If opt5.Value = False Then TBGravar!opt5_Nivel = False Else TBGravar!opt5_Nivel = True
        TBGravar.Update
        USMsgBox ("Nível cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/PPAP/PSW/Submissão/Nível"
        Evento = "Cadastrar nível"
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
