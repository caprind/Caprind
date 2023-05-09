VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_carteira_menuimpressao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Vendas - Follow up"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3360
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
   ScaleHeight     =   3045
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   661
      DibPicture      =   "frmVendas_carteira_menuImpressao.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_carteira_menuImpressao.frx":7180
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   2640
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Menu relatórios"
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
      Height          =   1935
      Left            =   150
      TabIndex        =   2
      Top             =   510
      Width           =   3075
      Begin DrawSuite2022.USButton Cmd_detalhado 
         Height          =   690
         Left            =   210
         TabIndex        =   0
         Top             =   300
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   1217
         DibPicture      =   "frmVendas_carteira_menuImpressao.frx":749A
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Detalhado"
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
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         PicAlign        =   7
         PicSize         =   1
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton Cmd_resumido 
         Height          =   660
         Left            =   210
         TabIndex        =   1
         Top             =   1080
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   1164
         DibPicture      =   "frmVendas_carteira_menuImpressao.frx":D77E
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         Caption         =   "Resumido"
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
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         PicAlign        =   7
         PicSize         =   1
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   5
      End
   End
End
Attribute VB_Name = "frmVendas_carteira_menuimpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_detalhado_Click()
On Error GoTo tratar_erro

NomeRel = "Vendas_follow up_detalhado.rpt"
ProcImprimirRel frmVendas_carteira.FormulaRel_Vendas_Carteira, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_resumido_Click()
On Error GoTo tratar_erro

NomeRel = "Vendas_follow up_resumido.rpt"
ProcImprimirRel frmVendas_carteira.FormulaRel_Vendas_Carteira, ""

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
