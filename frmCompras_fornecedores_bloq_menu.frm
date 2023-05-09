VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_fornecedores_bloq_menu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fornecedores - Desbloquear"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   55
      TabIndex        =   0
      Top             =   0
      Width           =   3285
      Begin DrawSuite2022.USButton Cmd_padrao 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Total"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2022.USButton cmdParcial 
         Height          =   360
         Left            =   180
         TabIndex        =   2
         Top             =   630
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Parcial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
   End
End
Attribute VB_Name = "frmCompras_fornecedores_bloq_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_padrao_Click()
On Error GoTo tratar_erro

With frmCompras_fornecedores_bloq
    If .txtstatus.Text = "Liberado" Then
        USMsgBox ("O fornecedor " & frmCompras_fornecedores.txtnomerazao.Text & " já esta liberado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    .txtstatus.Text = "Liberado"
    .txtresponsavel = pubUsuario
    .txtobservacoes.Text = ""
    .txtobservacoes.Locked = True
    .txtobservacoes.TabStop = False
    Unload Me
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdParcial_Click()
On Error GoTo tratar_erro

With frmCompras_fornecedores_bloq
    If .txtstatus.Text = "Parcial" Then
        USMsgBox ("O fornecedor " & frmCompras_fornecedores.txtnomerazao.Text & " já esta liberado parcialmente."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    .txtstatus.Text = "Parcial"
    .txtobservacoes.Text = ""
    .txtresponsavel = pubUsuario
    .txtobservacoes.Locked = True
    .txtobservacoes.TabStop = False
    Unload Me
End With

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

