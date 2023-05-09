VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMenuCaprind 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Menu Caprind v4.9"
   ClientHeight    =   9960
   ClientLeft      =   540
   ClientTop       =   300
   ClientWidth     =   15120
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmMenucaprind.frx":0000
   ScaleHeight     =   9960
   ScaleWidth      =   15120
   WindowState     =   2  'Maximizado
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8925
      Left            =   210
      TabIndex        =   0
      Top             =   630
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   15743
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   53
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5370
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   9
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenucaprind.frx":191DF
            Key             =   "A"
            Object.Tag             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenucaprind.frx":19549
            Key             =   "Menu"
            Object.Tag             =   "Menu"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenucaprind.frx":195C6
            Key             =   "B"
            Object.Tag             =   "B"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparente
      Caption         =   "www.caprind.com.br"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   13320
      TabIndex        =   1
      Top             =   9540
      Width           =   1470
   End
End
Attribute VB_Name = "frmMenuCaprind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaMenu TreeView1, ImageList1
Caption = "CÓPIA REGISTRADA"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo tratar_erro

ProcAbreModuloMenuTreeView (Node.key)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
