VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmestoque_item_atualizar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Estoque | Movimentação - Atualizar"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5490
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   4200
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   3795
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton btnAtualizar 
      Height          =   615
      Left            =   3390
      TabIndex        =   7
      Top             =   3030
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1085
      DibPicture      =   "frmestoque_item_atualizar.frx":0000
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Atualizar (F3)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   873
      DibPicture      =   "frmestoque_item_atualizar.frx":A1AD
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmestoque_item_atualizar.frx":1435A
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
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
      Height          =   1935
      Left            =   510
      TabIndex        =   5
      Top             =   870
      Width           =   4425
      Begin VB.CheckBox chkSaldoRE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "05 - Saldo das RE's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CheckBox Chk5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "05 - Empenho do RE"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1290
         Width           =   4095
      End
      Begin VB.CheckBox Chk4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "04 - RE de pedido com centro de custo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1020
         Width           =   4095
      End
      Begin VB.CheckBox Chk3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "03 - Valores de entrada (unitário e total)"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   765
         Width           =   4095
      End
      Begin VB.CheckBox Chk2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "02 - Custo de material nas ordens"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   495
         Width           =   4095
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "01 - Movimentação sem saldo total"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmestoque_item_atualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If Chk1.Value = 0 And Chk2.Value = 0 And Chk3.Value = 0 And Chk4.Value = 0 And Chk5.Value = 0 And chkSaldoRE = 0 Then
    USMsgBox ("Informe uma das opções antes de atualizar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
frmestoque_item.ProcAtualizacao
Unload Me

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

Private Sub btnAtualizar_Click()
On Error GoTo tratar_erro

   If USMsgBox("Deseja realmente fazer uma atualização de estoque?", vbYesNo, "CAPRIND v5.0") = vbYes Then
   ProcAtualizar
   End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcAtualizar
    Case 13: btnAtualizar_Click
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


