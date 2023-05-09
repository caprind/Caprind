VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmQualidade_RNC_Usuarios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RNC - Inserir usuário"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView2 
      Height          =   5880
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   10372
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuário"
         Object.Width           =   17639
      EndProperty
   End
End
Attribute VB_Name = "frmQualidade_RNC_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo tratar_erro

ListView2.ListItems.Clear

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select * from Usuarios Order By Nome", Conexao, adOpenKeyset, adLockOptimistic
Do While TBUsuarios.EOF = False
    ListView2.ListItems.Add = TBUsuarios!Nome
    TBUsuarios.MoveNext
Loop
TBUsuarios.Close



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub


Private Sub ListView2_DblClick()
On Error GoTo tratar_erro

With frmQualidade_RNC
If .RespRNC = "auditor" Then .txtAuditor = ListView2.SelectedItem.Text Else .txtRespQualidade = ListView2.SelectedItem.Text
Unload Me
    
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
