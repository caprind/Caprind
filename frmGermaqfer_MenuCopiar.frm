VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmGermaqfer_MenuCopiar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Copiar"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame2 
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
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3465
      Begin DrawSuite2022.USButton cmdPosto 
         Height          =   480
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   847
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Copiar turnos para outro posto de trabalho"
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
      Begin DrawSuite2022.USButton cmdTurno 
         Height          =   480
         Left            =   180
         TabIndex        =   2
         Top             =   840
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   847
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Copiar turno selecionado para o mesmo posto de trabalho"
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
Attribute VB_Name = "frmGermaqfer_MenuCopiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPosto_Click()
On Error GoTo tratar_erro

Sit_REG = 1
frmGermaqfer_Copiar.Show 1
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdTurno_Click()
On Error GoTo tratar_erro

With frmGermaqfer
    If .cmbdia = "" Then
        USMsgBox ("Informe o dia da semana antes de copiar."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If .Novo_maquina1 = True Then
        USMsgBox ("Salve o turno antes de copiar."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End With

Sit_REG = 2
frmGermaqfer_Copiar.Show 1
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

