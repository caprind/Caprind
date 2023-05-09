VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdministrador 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caprind - Liberação de licença de uso"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdministrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdministrador.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdministrador.frx":0970
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdministrador.frx":0E83
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdministrador.frx":181D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdministrador.frx":2110
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
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
      Height          =   2175
      Left            =   55
      TabIndex        =   3
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Txt_codigo_liberacao 
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
         Left            =   3885
         MaxLength       =   16
         MouseIcon       =   "frmAdministrador.frx":2E2C
         TabIndex        =   1
         ToolTipText     =   "Código de liberação."
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox Txt_codigo_seguranca 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3885
         Locked          =   -1  'True
         MouseIcon       =   "frmAdministrador.frx":3136
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código de segurança."
         Top             =   300
         Width           =   2295
      End
      Begin DrawSuite2014.USButton cmdliberacao 
         Height          =   855
         Left            =   3525
         TabIndex        =   2
         Top             =   1140
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1508
         BorderColor     =   8421504
         BorderColorDisabled=   0
         Caption         =   "Liberar licença de uso"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicNormal       =   "frmAdministrador.frx":3440
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin VB.Image Image1 
         Height          =   1905
         Left            =   180
         Picture         =   "frmAdministrador.frx":511A
         Top             =   180
         Width           =   1830
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de liberação :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2055
         TabIndex        =   5
         Top             =   660
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de segurança :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1935
         TabIndex        =   4
         Top             =   300
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmAdministrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdliberacao_Click()
On Error GoTo tratar_erro

If Txt_codigo_liberacao.Text = "" Then
    MsgBox ("Informe o código de liberação antes de liberar."), vbInformation
    Txt_codigo_liberacao.SetFocus
    Exit Sub
End If
frmLogon.ActiveLock1.LiberationKey = Txt_codigo_liberacao
Instalado = "SIM"
'Check if it was correct:
If Not (frmLogon.ActiveLock1.RegisteredUser) Then
    MsgBox "Código de liberação inválido.", vbInformation
    Txt_codigo_liberacao.SetFocus
Else
    MsgBox "O Caprind foi liberado para utilização. O sistema será fechado automaticamente, favor reiniciar.", vbInformation
    SaveSetting "Procam", "CaprindSQL", "Instalacao", Instalado
    SaveSetting "Procam", "CaprindSQL", "Nome", Nome
    frmabertura.Timer1.Enabled = True
    End
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Instalado = GetSetting("Procam", "CaprindSQL", "Instalacao")
Txt_codigo_seguranca = frmLogon.ActiveLock1.SoftwareCode
Txt_codigo_liberacao = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
