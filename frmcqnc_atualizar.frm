VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmcqnc_atualizar 
   Appearance      =   0  'Flat
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Qualidade - N�o conformidade - Atualizar"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5100
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
   Icon            =   "frmcqnc_atualizar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   5085
      _ExtentX        =   8969
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
      ButtonCaption1  =   "Atualizar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Utilizado pelo administrador do sistema."
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
      ButtonWidth1    =   50
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
      ButtonLeft2     =   54
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
      ButtonLeft3     =   58
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
      ButtonLeft4     =   96
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
      ButtonLeft5     =   124
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3780
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmcqnc_atualizar.frx":000C
         Count           =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   90
      TabIndex        =   3
      Top             =   1560
      Width           =   4875
      Begin VB.CheckBox Chk4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "04 - Sincronizar ordens com Ordens de servi�o"
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
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   4545
      End
      Begin VB.CheckBox Chk3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "03 - Status da(s) NC"
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
         Left            =   180
         TabIndex        =   2
         Top             =   706
         Width           =   4545
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "01 - ID do apontamento e dados da(s) NC"
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
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   4545
      End
      Begin VB.CheckBox Chk2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "02 - Qtde.(s) da(s) NC na(s) Ordem(ns) e OS('s)"
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
         Left            =   180
         TabIndex        =   1
         Top             =   443
         Width           =   4545
      End
   End
End
Attribute VB_Name = "frmcqnc_atualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If Chk1.Value = 0 And Chk2.Value = 0 And Chk3.Value = 0 And Chk4.Value = 0 Then
    USMsgBox ("Informe uma das op��es antes de atualizar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmcqnc.ProcAtualizacao
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcAtualizar
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4875, 5, True

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAtualizar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
