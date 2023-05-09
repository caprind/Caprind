VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmManutencaoOpcoes 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Manutenção - Menu opções"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   Icon            =   "frmManutencaoOpcoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Centralizar no Mestre
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   767
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   210
      TabIndex        =   0
      Top             =   570
      Width           =   3255
      Begin DrawSuite2022.USButton Btn_manutencao 
         Height          =   855
         Left            =   270
         TabIndex        =   1
         Top             =   1170
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         DibPicture      =   "frmManutencaoOpcoes.frx":000C
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Manutenção"
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
         PicAlign        =   8
         PicSize         =   5
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton btn_solicitacao 
         Height          =   855
         Left            =   270
         TabIndex        =   2
         Top             =   210
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         DibPicture      =   "frmManutencaoOpcoes.frx":11971
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         Caption         =   "Solicitação de manutenção"
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
         PicAlign        =   8
         PicSize         =   5
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   5
      End
   End
End
Attribute VB_Name = "frmManutencaoOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_manutencao_Click()
On Error GoTo tratar_erro

Unload Me

frmManutencao_menu.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btn_solicitacao_Click()
On Error GoTo tratar_erro

With frmManutencao
    .txttipo.Text = "Solicitação"
    .ProcHabilitarSolicitacao
    ProcCriaCodigoSolicitacao
End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriaCodigoSolicitacao()
On Error GoTo tratar_erro
Dim CodigoSol As String
Var = "S"

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao order by CodSOL", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast

CodSol = TBLISTA!CodSol
CodSol = Right(CodSol, 9)
CodSol = Left(CodSol, 6)

CodSol = Int(CodSol) + 1
    Select Case Len(CodSol)
        Case 1: CodigoSol = "00000" & CodSol
        Case 2: CodigoSol = "0000" & CodSol
        Case 3: CodigoSol = "000" & CodSol
        Case 4: CodigoSol = "00" & CodSol
        Case 5: CodigoSol = "0" & CodSol
    End Select
    Ano = Right(Year(Date), 2)
CodigoSol = "SOL-" & CodigoSol & "/" & Right(Year(Date), 2)
Else
    CodigoSol = "SOL-000001" & "/" & Right(Year(Date), 2)
End If
TBLISTA.Close
frmManutencao.txtCodigo.Text = CodigoSol

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

