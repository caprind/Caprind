VERSION 5.00
Begin VB.Form frmCompras_Pedido_NecessSolici 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de necessidades/solicitações"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   6450
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame2 
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
      Height          =   6375
      Left            =   55
      TabIndex        =   0
      Top             =   30
      Width           =   8745
      Begin VB.TextBox txtTexto 
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
         Height          =   5835
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Lista de necessidades/solicitações."
         Top             =   390
         Width           =   8355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Necessidades/Solicitações"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   63
         Left            =   3427
         TabIndex        =   2
         Top             =   180
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmCompras_Pedido_NecessSolici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select NecessSolici from Compras_pedido_lista where IdLista = " & frmCompras_Pedido.TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtTexto = IIf(IsNull(TBLISTA!NecessSolici), "", TBLISTA!NecessSolici)
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    'Case vbKeyF1: cmdAjuda_Click
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


