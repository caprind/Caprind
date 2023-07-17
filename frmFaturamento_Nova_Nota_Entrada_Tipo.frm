VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Nova_Nota_Entrada_Tipo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Finalidade entrada"
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_Nova_Nota_Entrada_Tipo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnContinuar 
      Height          =   285
      Left            =   2370
      TabIndex        =   6
      Top             =   1800
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   503
      Caption         =   "Continuar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   5249536
      BorderColorOver =   8076800
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   2205
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   714
      DibPicture      =   "frmFaturamento_Nova_Nota_Entrada_Tipo.frx":000C
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
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   270
      TabIndex        =   0
      Top             =   540
      Width           =   3225
      Begin DrawSuite2022.USOptionButton OptIndustrializacao 
         Height          =   255
         Left            =   330
         TabIndex        =   3
         Top             =   210
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Entrada para industrialização"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         Value           =   -1  'True
      End
      Begin DrawSuite2022.USOptionButton optConserto 
         Height          =   255
         Left            =   330
         TabIndex        =   4
         Top             =   480
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Entrada para conserto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USOptionButton optDevolucao 
         Height          =   255
         Left            =   330
         TabIndex        =   5
         Top             =   750
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Entrada de devolução de venda"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
      End
   End
End
Attribute VB_Name = "frmFaturamento_Nova_Nota_Entrada_Tipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnContinuar_Click()
On Error GoTo tratar_erro

Compras = OptIndustrializacao.Value
Vendas = optConserto.Value
Vendas = optDevolucao.Value

If USMsgBox("A finalidade da nota de entrada está correta?", vbYesNo, "CAPRIND  v5.0") = vbYes Then
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Compras = OptIndustrializacao.Value
Vendas = optConserto.Value
Vendas = optDevolucao.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optConserto_Click()
On Error GoTo tratar_erro

Compras = OptIndustrializacao.Value
Vendas = optConserto.Value
Vendas = optDevolucao.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub optDevolucao_Click()
On Error GoTo tratar_erro

Compras = OptIndustrializacao.Value
Vendas = optConserto.Value
Vendas = optDevolucao.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub OptIndustrializacao_Click()
On Error GoTo tratar_erro

Compras = OptIndustrializacao.Value
Vendas = optConserto.Value
Vendas = optDevolucao.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
