VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmvendas_tele_clientes_data 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alterar data"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   55
      TabIndex        =   2
      Top             =   -60
      Width           =   1935
      Begin VB.CommandButton CmdSalvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Picture         =   "frmvendas_tele_clientes_data.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salvar data do próximo contato (F3)"
         Top             =   390
         Width           =   315
      End
      Begin MSComCtl2.DTPicker Txtproximo 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Data do próximo contato."
         Top             =   390
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   130482177
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Próx. contato"
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
         Height          =   165
         Left            =   180
         TabIndex        =   3
         Top             =   180
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmvendas_tele_clientes_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdSalvar_Click()
On Error GoTo tratar_erro

Permitido = False
With frmVendas_Tele_Clientes.ListView1
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Conexao.Execute "UPDATE Vendas_Tele Set Proximo = '" & Txtproximo & "' where Codigo = " & .ListItems(InitFor).ListSubItems(7)
            '==================================
            Modulo = "Vendas/Clientes"
            Evento = "Alterar data do próximo contato"
            ID_documento = .ListItems(InitFor).ListSubItems(7)
            Documento = "Cliente: " & .ListItems(InitFor).ListSubItems(1) & " - Cidade: " & .ListItems(InitFor).ListSubItems(2)
            Documento1 = "Data: " & Txtproximo & " - Histórico: " & .ListItems(InitFor).ListSubItems(6)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
USMsgBox ("Alteração(ões) efetuada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
frmVendas_Tele_Clientes.ProcCarregaLista
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Txtproximo.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: CmdSalvar_Click
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
