VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmSenha 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Senha de liberação"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Senha de liberação"
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
      Height          =   1155
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4545
      Begin DrawSuite2022.USButton btnLiberar 
         Height          =   525
         Left            =   3090
         TabIndex        =   3
         ToolTipText     =   "Liberar alteração"
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         DibPicture      =   "frmSenha.frx":0000
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Liberar"
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
      _ExtentY        =   767
      DibPicture      =   "frmSenha.frx":7E58
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
      Icon            =   "frmSenha.frx":FCB0
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLiberar_Click()
On Error GoTo tratar_erro

If Formulario = "Estoque/Movimentação" Then
   If txtSenha.Text = "280362E" Then
    LiberarAlteracao = True
    Unload Me
   Else
    USMsgBox "Senha incorreta, tente novamente!", vbCritical, "CAPRIND v5.0"
    txtSenha.Text = ""
    txtSenha.SetFocus
    LiberarAlteracao = False
    Exit Sub
   End If
End If


If Formulario = "PCP/Gerenciamento de ordem" Then
   If txtSenha.Text = "280362O" Then
    LiberarAlteracao = True
    frmprod_atualizar.Show 1
    Unload Me
   Else
    USMsgBox "Senha incorreta, tente novamente!", vbCritical, "CAPRIND v5.0"
    txtSenha.Text = ""
    txtSenha.SetFocus
    LiberarAlteracao = False
    Exit Sub
   End If
End If

If Formulario = "Vendas/Pedido interno" Then
   If txtSenha.Text = "280362PI" Then
    LiberarAlteracao = True
    Unload Me
   Else
    USMsgBox "Senha incorreta, tente novamente!", vbCritical, "CAPRIND v5.0"
    txtSenha.Text = ""
    txtSenha.SetFocus
    LiberarAlteracao = False
    Exit Sub
   End If
End If

If Formulario = "Compras/Pedido" Or Formulario = "Compras/Pedido/Aprovar" Then
   If txtSenha.Text = "280362PC" Then
    LiberarAlteracao = True
    Unload Me
   Else
    USMsgBox "Senha incorreta, tente novamente!", vbCritical, "CAPRIND v5.0"
    txtSenha.Text = ""
    txtSenha.SetFocus
    LiberarAlteracao = False
    Exit Sub
   End If
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
  Case 13: btnLiberar_Click
  Case vbKeyEscape: Unload Me
End Select
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

