VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProcessos_bloq 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Engenharia - Processos - Status"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4860
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmProcessos_bloq.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3015
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
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   0
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
         Height          =   2025
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
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
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   1
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   42
      ButtonTop2      =   2
      ButtonWidth2    =   50
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   94
      ButtonTop3      =   2
      ButtonWidth3    =   68
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
      ButtonLeft4     =   164
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   168
      ButtonTop5      =   2
      ButtonWidth5    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   206
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   234
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
      ButtonUseMaskColor7=   0   'False
   End
End
Attribute VB_Name = "frmProcessos_bloq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcBloquear()
On Error GoTo tratar_erro

If txtStatus.Text = "Bloqueado" Then
    USMsgBox ("O processo " & frmProcessos.txtidprocesso.Text & " já está bloqueado."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
txtObservacoes.Locked = False
txtObservacoes.Text = ""
txtStatus.Text = "Bloqueado"
txtObservacoes.SetFocus
txtResponsavel.Text = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcDesbloquear()
On Error GoTo tratar_erro

If txtStatus.Text = "Liberado" Then
    USMsgBox ("O processo " & frmProcessos.txtidprocesso.Text & " já está liberado."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
txtObservacoes.Locked = False
txtObservacoes.Text = ""
txtStatus.Text = "Liberado"
txtObservacoes.SetFocus
txtResponsavel = pubUsuario

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

With frmProcessos
    If txtObservacoes = "" Then
        USMsgBox ("Informe a observação antes de salvar."), vbExclamation, "CAPRIND v5.0"
        If txtObservacoes.Enabled = True Then txtObservacoes.SetFocus
        Exit Sub
    End If
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select * from processos where IDProcesso = " & .txtidprocesso.Text, Conexao, adOpenKeyset, adLockOptimistic
    If txtStatus.Text = "Liberado" Then
       TBProcessos!Bloqueado = False
    Else
       TBProcessos!Bloqueado = True
    End If
    TBProcessos!obs_Status = txtObservacoes.Text
    TBProcessos!resp_Status = txtResponsavel.Text
    TBProcessos.Update
    If txtStatus.Text = "Liberado" Then
        USMsgBox ("Processo " & .txtidprocesso.Text & " liberado com sucesso."), vbInformation, "CAPRIND v5.0"
    ElseIf txtStatus.Text = "Bloqueado" Then
        USMsgBox ("Processo " & .txtidprocesso.Text & " bloqueado com sucesso."), vbInformation, "CAPRIND v5.0"
    End If
    txtObservacoes.Locked = True
    .txtStatus.Text = txtStatus.Text
    .ProcCarregaListaProcessos (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
End With
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
    Case vbKeyF3: ProcSalvar
    Case vbKeyF6: ProcBloquear
    Case vbKeyF7: ProcDesbloquear
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 6330, 7, True
Set TBProcessos = CreateObject("adodb.recordset")
TBProcessos.Open "Select * from processos where IDProcesso = " & frmProcessos.txtidprocesso.Text, Conexao, adOpenKeyset, adLockOptimistic
If TBProcessos.EOF = False Then
    If TBProcessos!Bloqueado = True Then
        txtStatus.Text = "Bloqueado"
    Else
        txtStatus.Text = "Liberado"
    End If
    txtObservacoes.Text = IIf(IsNull(TBProcessos!obs_Status), "", TBProcessos!obs_Status)
    txtResponsavel.Text = IIf(IsNull(TBProcessos!resp_Status), "", TBProcessos!resp_Status)
End If
TBProcessos.Close

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

