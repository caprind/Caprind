VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_proposta_Comissao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Vendedor | % Comissão"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   3180
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
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
      Height          =   885
      Left            =   630
      TabIndex        =   2
      Top             =   990
      Width           =   1815
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   1470
         Top             =   330
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   344
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
      Begin VB.TextBox txtComissao 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   270
         Width           =   1065
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   820
      DibPicture      =   "frmVendas_proposta_Comissao.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_proposta_Comissao.frx":A123
      ShowControlBox  =   0   'False
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   240
      Left            =   300
      Top             =   570
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   423
      Caption         =   "Informe o % da sua comissão"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      NoHTMLCaption   =   "Informe o % da sua comissão"
   End
   Begin DrawSuite2022.USButton btnAceitar 
      Height          =   675
      Left            =   630
      TabIndex        =   3
      ToolTipText     =   "Carregar o valor da comissão no cliente."
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1191
      DibPicture      =   "frmVendas_proposta_Comissao.frx":A43D
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Aceitar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      PicAlign        =   3
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      Theme           =   4
   End
End
Attribute VB_Name = "frmVendas_proposta_Comissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAceitar_Click()
On Error GoTo tratar_erro

If txtComissao.Text = "" Then
USMsgBox "Informe o valor da comissão para esse cliente"
txtComissao.SetFocus
Exit Sub
Else
Comissao = txtComissao.Text
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtComissao_LostFocus()
On Error GoTo tratar_erro

txtComissao.Text = Format(txtComissao.Text, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
