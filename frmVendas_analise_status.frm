VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_analise_status 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Análise crítica - Status"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   2775
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   55
      TabIndex        =   9
      Top             =   -30
      Width           =   2775
      Begin VB.OptionButton OPT6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DECLINADA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2200
      End
      Begin VB.OptionButton OPT1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ABERTA EM ANALISE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2200
      End
      Begin VB.OptionButton OPT3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CANCELADA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2200
      End
      Begin VB.OptionButton OPT4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PERDIDA P/ PRAZO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2200
      End
      Begin VB.OptionButton OPT5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PERDIDA P/ PREÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2200
      End
      Begin VB.OptionButton OPT2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "APROVADA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2200
      End
   End
   Begin DrawSuite2022.USButton Cmd_OK 
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   2370
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "OK (F3)"
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
      Theme           =   1
   End
   Begin DrawSuite2022.USButton Cmd_sair 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2370
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "Sair (Esc)"
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
      Theme           =   1
   End
   Begin DrawSuite2022.USButton Cmd_ajuda 
      Height          =   375
      Left            =   885
      TabIndex        =   7
      Top             =   2370
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "Ajuda (F1)"
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
      Theme           =   1
   End
End
Attribute VB_Name = "frmVendas_analise_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_OK_Click()
On Error GoTo tratar_erro

With frmVendas_analise
    If OPT1.Value = True Then
        .Txt_status = OPT1.Caption
        .ProcEscondeDataStatus
    ElseIf OPT2.Value = True Then
            .Txt_status = OPT2.Caption
            .ProcMostraDataStatus
        ElseIf OPT3.Value = True Then
                .Txt_status = OPT3.Caption
                .ProcMostraDataStatus
            ElseIf OPT4.Value = True Then
                    .Txt_status = OPT4.Caption
                    .ProcMostraDataStatus
                ElseIf OPT5.Value = True Then
                        .Txt_status = OPT5.Caption
                        .ProcMostraDataStatus
                    ElseIf opt6.Value = True Then
                            .Txt_status = opt6.Caption
                            .ProcMostraDataStatus
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: Cmd_OK_Click
    'Case vbKeyF1: Cmd_ajuda_Click
    Case vbKeyEscape: Cmd_sair_Click
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_sair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
