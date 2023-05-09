VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_recebimento_Menuimpressao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu relatórios"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2655
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
      Height          =   1875
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   2565
      Begin DrawSuite2022.USButton Cmd_inspecao_recebimento 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Inspeção de recebimento"
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
      Begin DrawSuite2022.USButton Cmd_etiqueta 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   690
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Etiqueta de identificação"
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
      Begin DrawSuite2022.USButton Cmd_etiqueta_personalizada 
         Height          =   570
         Left            =   180
         TabIndex        =   2
         Top             =   1140
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   1005
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Etiqueta de identificação personalizada"
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
Attribute VB_Name = "frmCompras_recebimento_Menuimpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_etiqueta_personalizada_Click()
On Error GoTo tratar_erro

ProcAbreModuloEtiqueta True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_etiqueta_Click()
On Error GoTo tratar_erro

ProcAbreModuloEtiqueta False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_inspecao_recebimento_Click()
On Error GoTo tratar_erro

If VerifCampos = False Then Exit Sub
NomeRel = "CQ_inspecao recebimento.rpt"
ProcImprimirRel "{Compras_recebimento.ID} = " & frmCompras_recebimento.txtid, ""

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

Private Sub ProcAbreModuloEtiqueta(Personalizado As Boolean)
On Error GoTo tratar_erro

If VerifCampos = False Then Exit Sub
Inspecao_recebimento = True
Estoque_recebimento = False
Faturamento = False
Permitido = Personalizado
frmestoque_item_imprimir_etiqueta.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function VerifCampos() As Boolean
On Error GoTo tratar_erro

VerifCampos = True
With frmCompras_recebimento
    If .Txt_lote = "" Then
        USMsgBox ("Informe o número do lote antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
        VerifCampos = False
        Exit Function
    End If
    If .ListProdReceb.ListItems.Count = 0 Then
        USMsgBox ("É necessário ter produtos inspecionados antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
        VerifCampos = False
        Exit Function
    End If
End With

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
