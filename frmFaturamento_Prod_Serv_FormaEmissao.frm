VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_FormaEmissao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Forma de emissão"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1155
      Left            =   55
      TabIndex        =   2
      Top             =   0
      Width           =   2355
      Begin DrawSuite2014.USButton cmdCaprind 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   210
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Caprind"
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
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2014.USButton cmdGNFe 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "GNFe"
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
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_FormaEmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCaprind_Click()
On Error GoTo tratar_erro

frmFaturamento_Prod_Serv_NFe.Show
Unload Me
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub cmdGNFe_Click()
On Error GoTo tratar_erro

frmFaturamento_Prod_Serv_Migrate.Show
Unload Me
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub
