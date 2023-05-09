VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_propostaII_Status 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Status"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2895
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   2895
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
      Height          =   2595
      Left            =   55
      TabIndex        =   5
      Top             =   0
      Width           =   2775
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
         Height          =   435
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2200
      End
      Begin VB.OptionButton OPT2 
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
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Top             =   690
         Width           =   2200
      End
      Begin VB.OptionButton OPT3 
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
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   1140
         Width           =   2200
      End
      Begin VB.OptionButton OPT4 
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
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1590
         Width           =   2200
      End
      Begin VB.OptionButton OPT5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PORTAL ELETRONICO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   2200
      End
   End
   Begin DrawSuite2022.USButton cmdfechar 
      Height          =   360
      Left            =   60
      TabIndex        =   6
      Top             =   2610
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "OK"
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
Attribute VB_Name = "frmVendas_propostaII_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFechar_Click()
On Error GoTo tratar_erro

With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
    If OPT1.Value = True Then
        .txtstatus.Text = OPT1.Caption
        .txt_datamodificado.Text = ""
    ElseIf OPT2.Value = True Then
            .txtstatus.Text = OPT2.Caption
        ElseIf OPT3.Value = True Then
                .txtstatus.Text = OPT3.Caption
            ElseIf OPT4.Value = True Then
                    .txtstatus.Text = OPT4.Caption
                Else
                    .txtstatus.Text = OPT5.Caption
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
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
