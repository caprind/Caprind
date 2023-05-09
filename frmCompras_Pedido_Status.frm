VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_Pedido_Status 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Compras Pedido | Alterar Status"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4260
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
   ScaleHeight     =   3600
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alterar Status"
      Height          =   2265
      Left            =   300
      TabIndex        =   2
      Top             =   630
      Width           =   3735
      Begin DrawSuite2022.USButton cmd_Aprovado 
         Height          =   1590
         Left            =   780
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Voltar o status do pedido para ""Aprovado""."
         Top             =   390
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   2805
         DibPicture      =   "frmCompras_Pedido_Status.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Voltar status para ""APROVADO"""
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
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton Cmd_Comprado 
         Height          =   1590
         Left            =   750
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Mudar o status do pedido para ""Comprado""."
         Top             =   390
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2805
         DibPicture      =   "frmCompras_Pedido_Status.frx":16B6D
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Mudar status para ""COMPRADO"""
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
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   3195
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   767
      DibPicture      =   "frmCompras_Pedido_Status.frx":1DCED
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
      Icon            =   "frmCompras_Pedido_Status.frx":20013
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmCompras_Pedido_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_OK_Click()
On Error GoTo tratar_erro

With frmCompras_Pedido
    If opt1.Value = True Then .txtStatus = opt1.Caption Else .txtStatus = opt2.Caption
    If .txtStatus.Text = "COMPRADO" Then
    Conexao.Execute "UPDATE Compras_pedido Set Status_Pedido = 'ABERTO' where IDpedido = " & .txtIDPedido
    Conexao.Execute "UPDATE Compras_pedido_Lista Set Status_item = 'N_RECEBIDO' where IDpedido = " & .txtIDPedido
    
    USMsgBox "Status modificado com sucesso!", vbInformation, "CAPRIND v5.0"
    .ProcAtualizalistapedido (1)
    End If
    
    If .txtStatus.Text = "APROVADO" Then
    Conexao.Execute "UPDATE Compras_pedido Set Status_Pedido = 'APROVADO' where IDpedido = " & .txtIDPedido
    Conexao.Execute "UPDATE Compras_pedido_Lista Set Status_item = 'APROVADO' where IDpedido = " & .txtIDPedido
    
    USMsgBox "Status do pedido modificado com sucesso!", vbInformation, "CAPRIND v5.0"
    .ProcAtualizalistapedido (1)
    End If
    
End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmd_Aprovado_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente alterar o status do pedido de compras para APROVADO?", vbYesNo) = vbYes Then
With frmCompras_Pedido
    If .txtStatus.Text = "COMPRADO" Then
    Conexao.Execute "UPDATE Compras_pedido Set Status_Pedido = 'APROVADO' where IDpedido = " & .txtIDPedido
    Conexao.Execute "UPDATE Compras_pedido_Lista Set Status_item = 'APROVADO' where IDpedido = " & .txtIDPedido
    
    USMsgBox "Status do pedido modificado com sucesso!", vbInformation, "CAPRIND v5.0"
    .txtStatus.Text = "APROVADO"
    .ProcAtualizalistapedido (1)
    End If
End With
End If

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_Comprado_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente alterar o status do pedido de compras para COMPRADO?", vbYesNo) = vbYes Then
With frmCompras_Pedido
    If .txtStatus.Text = "APROVADO" Then
    Conexao.Execute "UPDATE Compras_pedido Set Status_Pedido = 'ABERTO' where IDpedido = " & .txtIDPedido
    Conexao.Execute "UPDATE Compras_pedido_Lista Set Status_item = 'N_RECEBIDO' where IDpedido = " & .txtIDPedido
    
    USMsgBox "Status modificado com sucesso!", vbInformation, "CAPRIND v5.0"
    .txtStatus.Text = "COMPRADO"
    .ProcAtualizalistapedido (1)
    End If
End With
End If

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

Private Sub Form_Load()
On Error GoTo tratar_erro

With frmCompras_Pedido
    If .txtStatus.Text = "APROVADO" Then
    cmd_Aprovado.Visible = False
    Cmd_Comprado.Visible = True
    Else
    cmd_Aprovado.Visible = True
    Cmd_Comprado.Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
