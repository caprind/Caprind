VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form Frmvencimento 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   2370
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView calendario 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   156958721
      TitleBackColor  =   8421504
      TitleForeColor  =   16777215
      TrailingForeColor=   255
      CurrentDate     =   39059
   End
End
Attribute VB_Name = "Frmvencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calendario_DateClick(ByVal DateClicked As Date)
On Error GoTo tratar_erro

frmEstoque_Recebimento.txtVencimento.Text = Format(DateClicked, "dd/mm/yyyy")

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

calendario.Value = Date
FrmCalendario.Caption = "Calendário de " & Format(Date, "yyyy")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
