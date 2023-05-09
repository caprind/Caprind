VERSION 5.00
Begin VB.Form frmLogoff 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2370
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Timer Timer2 
      Interval        =   4000
      Left            =   2580
      Top             =   1950
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2160
      Top             =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparente
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2970
      TabIndex        =   1
      Top             =   2550
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparente
      Caption         =   "Conexão com o banco sendo fechada..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   180
      TabIndex        =   0
      Top             =   1350
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1200
      Picture         =   "frmLogoff.frx":0000
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frmLogoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

Label1.Left = 600
Label1.Caption = "Efetuando logoff do sistema..."

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Timer2_Timer()
On Error GoTo tratar_erro

Logoff = True
Unload Me
Unload frmMenucaprind_menulateral
Unload frmMDI
frmabertura.Timer2.Enabled = True
frmabertura.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
