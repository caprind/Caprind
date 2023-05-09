VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmManutencao_menu 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'Nenhum
   Caption         =   "  Manutenção - Menu"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4650
   Icon            =   "frmManutencao_menu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   661
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo do serviço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      TabIndex        =   7
      Top             =   2550
      Width           =   4485
      Begin VB.CheckBox chkHidraulica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hidraulica"
         Height          =   255
         Left            =   2310
         TabIndex        =   12
         Top             =   300
         Width           =   1065
      End
      Begin VB.CheckBox chkOutros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outros"
         Height          =   255
         Left            =   3450
         TabIndex        =   10
         Top             =   300
         Width           =   825
      End
      Begin VB.CheckBox chkMecanica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mecânica"
         Height          =   255
         Left            =   1110
         TabIndex        =   9
         Top             =   300
         Width           =   1065
      End
      Begin VB.CheckBox chkeletrica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Elétrica"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nova manutenção"
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
      Height          =   555
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   4485
      Begin VB.OptionButton optPredial 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Predial"
         Height          =   195
         Left            =   1860
         TabIndex        =   11
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optPosto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipamento"
         DisabledPicture =   "frmManutencao_menu.frx":000C
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optProduto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produto final"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3060
         TabIndex        =   5
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   4485
      Begin DrawSuite2022.USButton cmdCorretiva 
         Height          =   855
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1508
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Corretiva"
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
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   4
      End
      Begin DrawSuite2022.USButton cmdPreventiva 
         Height          =   855
         Left            =   1545
         TabIndex        =   2
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1508
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Preventiva"
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
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   3
      End
      Begin DrawSuite2022.USButton cmdPreditiva 
         Height          =   855
         Left            =   3030
         TabIndex        =   1
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1508
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         Caption         =   "Preditiva"
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
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   5
      End
   End
End
Attribute VB_Name = "frmManutencao_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCorretiva_Click()
On Error GoTo tratar_erro

If Optproduto.Value = True Then frmManutencao.Manutencao_Produto = True Else frmManutencao.Manutencao_Produto = False
frmManutencao.txttipo.Text = "Corretiva"
frmManutencao_Solicitacao_Abrir.Show 1
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPreditiva_Click()
On Error GoTo tratar_erro

With frmManutencao
    .txttipo.Text = "Preditiva"
    .ProcHabilitarPrevCorr
    If Optproduto.Value = True Then .Manutencao_Produto = True Else .Manutencao_Produto = False
    frmManutencao_Solicitacao_Abrir.Show 1
    'ProcCriaCodigoManutencao
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub cmdPreventiva_Click()
On Error GoTo tratar_erro

With frmManutencao
    .txttipo.Text = "Preventiva"
    .ProcHabilitarPrevCorr
    If Optproduto.Value = True Then .Manutencao_Produto = True Else .Manutencao_Produto = False
    frmManutencao_Solicitacao_Abrir.Show 1
    'ProcCriaCodigoManutencao
    
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriaCodigoManutencao()
On Error GoTo tratar_erro
Dim CodigoMan As String
Var = "S"

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from manutencao where Tipo <> '" & Var & "' order by CodMan", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast

CodigoMan = TBLISTA!codman
CodigoMan = Right(CodigoMan, 9)
CodigoMan = Left(CodigoMan, 6)

CodigoMan = Int(CodigoMan) + 1
    Select Case Len(CodigoMan)
        Case 1: CodigoMan = "00000" & CodigoMan
        Case 2: CodigoMan = "0000" & CodigoMan
        Case 3: CodigoMan = "000" & CodigoMan
        Case 4: CodigoMan = "00" & CodigoMan
        Case 5: CodigoMan = "0" & CodigoMan
    End Select
    Ano = Right(Year(Date), 2)
CodigoMan = "MAN-" & CodigoMan & "/" & Right(Year(Date), 2)
Else
    CodigoMan = "MAN-000001" & "/" & Right(Year(Date), 2)
End If
TBLISTA.Close
frmManutencao.txtCodigo.Text = CodigoMan

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

