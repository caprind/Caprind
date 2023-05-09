VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCQ_Certificado_Localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "CQ | Certificados | Localizar"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   4005
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por"
      Height          =   945
      Left            =   210
      TabIndex        =   4
      Top             =   570
      Width           =   3585
      Begin VB.OptionButton optCertificado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Certificado"
         Height          =   195
         Left            =   2250
         TabIndex        =   7
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton optLote 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lote"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Optproduto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto"
         Height          =   195
         Left            =   1110
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados para pesquisa"
      Height          =   2205
      Left            =   210
      TabIndex        =   1
      Top             =   1530
      Width           =   3585
      Begin DrawSuite2022.USButton btnLocalizar 
         Height          =   705
         Left            =   390
         TabIndex        =   3
         ToolTipText     =   "Clique aqui para localizar o certificado de analise"
         Top             =   1260
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1244
         DibPicture      =   "frmCQ_Certificado_Localizar.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Localizar"
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
         PicAlign        =   8
         Theme           =   4
      End
      Begin VB.TextBox txtDocumento 
         Height          =   345
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Documento para pesquisa"
         Top             =   750
         Width           =   2865
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Left            =   840
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   344
         Caption         =   "Documento para pesquisa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Documento para pesquisa"
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   714
      DibPicture      =   "frmCQ_Certificado_Localizar.frx":3650
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmCQ_Certificado_Localizar.frx":6CA0
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmCQ_Certificado_Localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLocalizar_Click()
On Error GoTo tratar_erro

If txtDocumento.Text = "" Then
USMsgBox "Digite os dados do documento para pesquisa", vbInformation, "CAPRIND v5.0"
txtDocumento.SetFocus
Exit Sub
End If

If optCertificado.Value = True Then
StrSql = "Select * from CQ_Certificado where CodCertificado = '" & txtDocumento.Text & "'"
End If

If optLote.Value = True Then
StrSql = "Select * from CQ_Certificado where Lote = '" & txtDocumento.Text & "'"
End If

If Optproduto.Value = True Then
StrSql = "Select * from CQ_Certificado where Produto = '" & txtDocumento.Text & "'"
End If

frmCQ_Certificado_Analise.ProcBuscaDadosLaudo

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
