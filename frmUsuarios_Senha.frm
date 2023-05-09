VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmUsuarios_Senha 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Usuários |  Senha de acesso ao sistema"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   1425
      Left            =   270
      Top             =   810
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2514
      Caption         =   "Atenção, sua senha é pessoal e instransferível, nunca informe a ninguém."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      NoHTMLCaption   =   $"frmUsuarios_Senha.frx":0000
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Digite sua senha para acesso ao sistema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1185
      Left            =   240
      TabIndex        =   1
      Top             =   2670
      Width           =   4545
      Begin DrawSuite2022.USButton btnGravarSenha 
         Height          =   525
         Left            =   3090
         TabIndex        =   3
         ToolTipText     =   "Gravar senha de acesso ao sistema"
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         DibPicture      =   "frmUsuarios_Senha.frx":0050
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Cadastrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   330
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Informe a senha para liberação"
         Top             =   360
         Width           =   2715
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   741
      DibPicture      =   "frmUsuarios_Senha.frx":7EA8
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmUsuarios_Senha.frx":FD00
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   1620
      Left            =   3000
      Top             =   690
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   2858
      Image           =   "frmUsuarios_Senha.frx":1001A
      Props           =   5
   End
End
Attribute VB_Name = "frmUsuarios_Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnGravarSenha_Click()
On Error GoTo tratar_erro
  
If txtSenha.Text <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from usuarios where usuario = '" & frmabertura.txtUsuario.Text & "' order by usuario", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir!Senha = txtSenha.Text
        TBAbrir.Update
    USMsgBox "Senha cadastrada com sucesso, utilize essa senha sempre que for usar o sistema.", vbInformation, "CAPRIND v5.0"
    End If
    TBAbrir.Close
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
'  Case 13: btnGravarSenha_Click
  Case vbKeyEscape: Unload Me
End Select
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
