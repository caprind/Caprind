VERSION 5.00
Begin VB.Form frmsimbolos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Símbolos"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.ComboBox cmbSimbolo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      ItemData        =   "frmsimbolos.frx":0000
      Left            =   105
      List            =   "frmsimbolos.frx":0002
      MouseIcon       =   "frmsimbolos.frx":0004
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Símbolos."
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmsimbolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'case vbkeyF1: cmdAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
    
cmbSimbolo.Clear
For i = 1 To Len(Simbolos)
    cmbSimbolo.AddItem Mid(Simbolos, i, 1)
Next

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbSimbolo_Click()
On Error GoTo tratar_erro

If RNC_Nao_Conformidade = True Then frmcqnc_retrabalho.Txt_instrucoes.SelText = cmbSimbolo.List(cmbSimbolo.ListIndex) Else frmProcessos.txtDescricao.SelText = cmbSimbolo.List(cmbSimbolo.ListIndex)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
