VERSION 5.00
Begin VB.Form frmFaturamento_Prod_Serv_NFSe_Log 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nota fiscal - Dados da NFSe - Log "
   ClientHeight    =   6345
   ClientLeft      =   1755
   ClientTop       =   1650
   ClientWidth     =   8160
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_Prod_Serv_NFSe_Log.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6345
      Index           =   17
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   8010
      Begin VB.CommandButton cmdDados_adicionais 
         BackColor       =   &H00C0C0C0&
         Height          =   2685
         Left            =   14670
         Picture         =   "frmFaturamento_Prod_Serv_NFSe_Log.frx":1042
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cadastrar/lLocalizar dados adicionais."
         Top             =   270
         Width           =   345
      End
      Begin VB.TextBox txtLog 
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
         Height          =   5955
         Left            =   180
         MaxLength       =   4800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Log de erros da nota fiscal."
         Top             =   270
         Width           =   7635
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_NFSe_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

If Sit_REG = 1 Then
    IDnota = frmFaturamento_Prod_Serv_NFSe.txtID_nota
    Caption = "Nota fiscal - Dados da NFSe - Log de erros"
Else
    'IDnota = frmFaturamento_Prod_Serv_NFe.txtID_nota
    'Caption = "Nota fiscal - Dados da NFe - Log de erros"
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select LogErro from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & IDnota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then txtLog = IIf(IsNull(TBAbrir!LogErro), "", TBAbrir!LogErro)
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
