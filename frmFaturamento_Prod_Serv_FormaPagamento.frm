VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_FormaPagamento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "NFe | Forma de pagamento"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtxPag 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   870
      Width           =   4755
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   1845
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_Prod_Serv_FormaPagamento.frx":0000
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
      Icon            =   "frmFaturamento_Prod_Serv_FormaPagamento.frx":A123
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USButton btnSalvar 
      Height          =   435
      Left            =   3690
      TabIndex        =   4
      ToolTipText     =   "Salvar alterações."
      Top             =   1320
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_Prod_Serv_FormaPagamento.frx":A43D
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "Salvar"
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
      PicSize         =   1
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição forma de pagamento (xPag - 50 caracteres)"
      Height          =   345
      Left            =   300
      TabIndex        =   3
      Top             =   630
      Width           =   4725
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_FormaPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalvar_Click()
On Error GoTo tratar_erro
  
Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
    End If
    TBGravar!Xpag = txtxPag.Text
    TBGravar.Update
    USMsgBox ("Descrição de forma de pagamento cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

txtxPag.Text = ""
  
Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
    txtxPag.Text = IIf(TBGravar!Xpag <> "", TBGravar!Xpag, "")
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"

End Sub
