VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frm_aguardando 
   BorderStyle     =   0  'Nenhum
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Centralizar no Mestre
   Begin DrawSuite2014.USAnimatedImage USAnimatedImage1 
      Height          =   5700
      Left            =   -1950
      TabIndex        =   0
      Top             =   -1620
      Width           =   6870
      _ExtentX        =   13441
      _ExtentY        =   10054
      GIF             =   "frm_aguardando.frx":0000
   End
End
Attribute VB_Name = "frm_aguardando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
'SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME
frmFaturamento_Prod_Serv_NFe_NS.ProcEnviarNotaSefaz
End Sub
