VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmproj_produto_MenuImpressao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu relatórios"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1965
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   1965
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
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   60
      TabIndex        =   5
      Top             =   1770
      Visible         =   0   'False
      Width           =   3225
      Begin VB.CommandButton Cmd_imprimir 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2715
         MouseIcon       =   "frmproj_produto_MenuImpressao.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmproj_produto_MenuImpressao.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Visualizar impressão (F5)"
         Top             =   390
         Width           =   315
      End
      Begin VB.ComboBox Cmb_nome_relatorio 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmproj_produto_MenuImpressao.frx":0248
         Left            =   180
         List            =   "frmproj_produto_MenuImpressao.frx":024A
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Opções de relatório."
         Top             =   390
         Width           =   2535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Nome do relatório"
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
         Height          =   195
         Left            =   690
         TabIndex        =   6
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1215
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   1875
      Begin DrawSuite2022.USButton Cmd_padrao 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Padrão"
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
      Begin DrawSuite2022.USButton Cmd_personalizado 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   690
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Personalizado"
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
Attribute VB_Name = "frmproj_produto_MenuImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_padrao_Click()
On Error GoTo tratar_erro

If Formulario = "Vendas/Produtos e serviços" Then
NomeRel = "Vendas_produtos_servicos.rpt"
End If

If Formulario = "Engenharia/Produtos e serviços" Then
NomeRel = "Engenharia_produtos_servicos.rpt"
End If

If Formulario = "Compras/Produtos e serviços" Then
NomeRel = "Compras_produtos_servicos.rpt"
End If

ProcImprimirRel frmproj_produto.FormulaRel_Produto, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_personalizado_Click()
On Error GoTo tratar_erro

NomeRel = "Engenharia_produtos_servicos_personalizado.rpt"
ProcImprimirRel frmproj_produto.FormulaRel_Produto, ""

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
