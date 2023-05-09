VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_PI 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Informações comerciais"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
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
   ScaleHeight     =   4530
   ScaleWidth      =   9015
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.TextBox txtCondicoes 
      Alignment       =   2  'Centralizar
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Condições de pagamento."
      Top             =   825
      Width           =   8625
   End
   Begin VB.TextBox txtobservacoes 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Observações."
      Top             =   1830
      Width           =   8655
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   4125
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   741
      DibPicture      =   "frmFaturamento_Prod_Serv_PI.frx":0000
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
      Icon            =   "frmFaturamento_Prod_Serv_PI.frx":7180
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Centralizar
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparente
      Caption         =   "Condições de pagamento"
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   6
      Left            =   3345
      TabIndex        =   5
      Top             =   570
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Centralizar
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparente
      Caption         =   "Observações"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   96
      Left            =   3465
      TabIndex        =   4
      Top             =   1590
      Width           =   2085
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_PI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo tratar_erro

ProcPuxaDadoscomerciais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDadoscomerciais()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM vendas_Proposta WHERE NCotacao = '" & frmFaturamento_Prod_Serv.txt_proposta.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

IDpedido = TBAbrir!Cotacao

'ProcLimparComercial
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM vendas_comercial WHERE cotacao = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
'    If TBCotacao!analize = "Sim" Or TBCotacao!analize = "Não" Then txtAnalize.Text = TBCotacao!analize
'    txtcalculos = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
'    txtimpostos = IIf(IsNull(TBCotacao!impostos), "", TBCotacao!impostos)
    txtCondicoes = IIf(IsNull(TBCotacao!condicoes), "", TBCotacao!condicoes)
'    txtgarantia = IIf(IsNull(TBCotacao!garantia), "", TBCotacao!garantia)
    txtobservacoes = IIf(IsNull(TBCotacao!observacoes), "", TBCotacao!observacoes)
'    txtReajuste = IIf(IsNull(TBCotacao!reajuste), "", TBCotacao!reajuste)
'     txttransporte = IIf(IsNull(TBCotacao!transporte), "", TBCotacao!transporte)
'     txtValidade = IIf(IsNull(TBCotacao!validade), "", TBCotacao!validade)
'     txtcalculos = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
End If
TBCotacao.Close
End If
TBAbrir.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

