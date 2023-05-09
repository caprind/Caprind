VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Estoque_Menu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Faturamento | Menu estoque"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
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
   ScaleHeight     =   5325
   ScaleWidth      =   4890
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escolha uma opção abaixo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   270
      TabIndex        =   1
      Top             =   720
      Width           =   4305
      Begin DrawSuite2022.USButton btnEstoqueSaldo 
         Height          =   495
         Left            =   390
         TabIndex        =   3
         Top             =   2670
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   873
         DibPicture      =   "frmFaturamento_Estoque_Menu.frx":0000
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Consultar saldo no estoque"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   741
      DibPicture      =   "frmFaturamento_Estoque_Menu.frx":1CAD
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
      Icon            =   "frmFaturamento_Estoque_Menu.frx":7F91
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmFaturamento_Estoque_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnConsultaMovimentacao_Click()
On Error GoTo tratar_erro

  frmFaturamento_Estoque_Movimentacao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub BtnConsultaSaldosNF_Click()
On Error GoTo tratar_erro

  frmFaturamento_Estoque.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnEstoqueSaldo_Click()
On Error GoTo tratar_erro

  frmFaturamento_Estoque_Saldos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub BtnExcluirMovimentacao_Click()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
ID_nota = .txtId
If ID_nota <> 0 Then
ApagarMovimentacaoNFe
USMsgBox "Movimentação excluida com sucesso!", vbInformation, "CAPRIND v5.0"
End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub BtnMovimentar_Click()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv
If USMsgBox("Deseja movimentar estoque com essa NFe?", vbYesNo, "CAPRIND v5.0") = vbYes Then
  If .txtId <> "" Then
  ID_empresa = .txtidempresa.Text
  ID_nota = .txtId.Text
    If ID_nota <> 0 And ID_empresa <> 0 Then
    '======================================
    ' Se for nota de saida baixa estoque
    '======================================
      If .Opt_entrada.Value = False Then
       BaixarEstoqueNF
       If Sair = True Then
       USMsgBox "Baixa executada com sucesso!", vbInformation, "CAPRIND v5.0"
       Else
       USMsgBox "Movimentação no estoque não executada, pois não existe(m) mais saldo(s) no(s) item(ns) da nota!", vbInformation, "CAPRIND v5.0"
       End If
       
      Else
       EntrarEstoqueNF
       USMsgBox "Entrada estoque executada com sucesso!", vbInformation, "CAPRIND v5.0"
      End If
    End If
  End If
End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

