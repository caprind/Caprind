VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmFaturamento_prod_Sefaz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NFe (SEFAZ)"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   5220
   Icon            =   "frmFaturamento_prod_Sefaz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5220
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5205
      Begin VB.OptionButton optNS 
         Caption         =   "Enviar NFe para a SEFAZ pelo CAPRIND"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.OptionButton optGNFe 
         Caption         =   "Enviar NFe para a SEFAZ pelo GNFe"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   510
         Width           =   3165
      End
      Begin VB.OptionButton optInvoicy 
         Caption         =   "Enviar NFe para a SEFAZ pelo Invoicy"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   3405
      End
      Begin DrawSuite2014.USButton cmdEnviarSefaz 
         Height          =   855
         Left            =   4080
         TabIndex        =   4
         ToolTipText     =   "Enviar nota fiscal eletrônica para a SEFAZ."
         Top             =   150
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1508
         DibPicture      =   "frmFaturamento_prod_Sefaz.frx":21F49
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Enviar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   8
         PicSize         =   5
         PicSizeH        =   32
         PicSizeW        =   32
         ToolTipTitle    =   "CAPRIND 5.0"
      End
   End
   Begin VB.Menu mnu_menu 
      Caption         =   "menu"
      Begin VB.Menu mnu_gravar 
         Caption         =   "Gravar"
      End
   End
End
Attribute VB_Name = "frmFaturamento_prod_Sefaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviarSefaz_Click()
On Error GoTo tratar_erro

If optNS.Value = True Then
frmFaturamento_Prod_Serv_NFe.Show
Unload Me
End If

If optGNFe.Value = True Then
frmFaturamento_Prod_Serv_Migrate.Show
Unload Me
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

 Set TBAbrir = CreateObject("adodb.recordset")
 TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
  If TBAbrir.EOF = False Then
   optGNFe.Value = TBAbrir!TPGNFe
   optNS.Value = TBAbrir!TPns
  End If
  TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub mnu_gravar_Click()
On Error GoTo tratar_erro
 
 Set TBAbrir = CreateObject("adodb.recordset")
 TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.Cmb_empresa.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
   TBAbrir!TPGNFe = optGNFe.Value
   TBAbrir!TPns = optNS.Value
   TBAbrir.Update
   USMsgBox "DAdos atualizados com sucesso!", vbInformation, "CAPRIND V5.0"
  End If
  TBAbrir.Close
  
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub
