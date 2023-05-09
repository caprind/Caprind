VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmestoque_item_imprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque | Relatórios"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   4830
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   4425
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   688
      DibPicture      =   "frmestoque_item_imprimir.frx":0000
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
      Icon            =   "frmestoque_item_imprimir.frx":1C95
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para impressão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   300
      TabIndex        =   3
      Top             =   630
      Width           =   4635
      Begin DrawSuite2022.USButton cmdEstoque 
         Height          =   870
         Left            =   270
         TabIndex        =   0
         ToolTipText     =   "Abrir menu de opções para relatórios do estoque"
         Top             =   450
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   1535
         DibPicture      =   "frmestoque_item_imprimir.frx":1FAF
         Caption         =   "Informações do estoque"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton cmdEtiqueta 
         Height          =   870
         Left            =   270
         TabIndex        =   1
         ToolTipText     =   "Imprimir etiqueta de identificação do item no estoque"
         Top             =   2250
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   1535
         DibPicture      =   "frmestoque_item_imprimir.frx":3C5C
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
         BorderColor     =   0
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   8421504
         GradientColor1  =   0
         GradientColor2  =   0
         GradientColor3  =   0
         GradientColor4  =   0
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   8421504
         GradientColorOver2=   8421504
         GradientColorOver3=   8421504
         GradientColorOver4=   8421504
         GradientColorDown1=   4210752
         GradientColorDown2=   4210752
         GradientColorDown3=   4210752
         GradientColorDown4=   4210752
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   6
      End
      Begin DrawSuite2022.USButton Cmd_identificacao_personalizada 
         Height          =   870
         Left            =   270
         TabIndex        =   2
         ToolTipText     =   "Imprimir etiqueta de identificação do item personalizada"
         Top             =   1350
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   1535
         DibPicture      =   "frmestoque_item_imprimir.frx":57B7D
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
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton btnVisualizar 
         Height          =   870
         Left            =   210
         TabIndex        =   5
         ToolTipText     =   "Visualizar relatório do estoque."
         Top             =   3510
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1535
         DibPicture      =   "frmestoque_item_imprimir.frx":593D1
         Caption         =   "Visualizar impressão"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
   End
End
Attribute VB_Name = "frmestoque_item_imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnVisualizar_Click()
On Error GoTo tratar_erro

NomeRel = "Estoque_saldo_resumido2.rpt"
'Debug.print FormulaRelatorio
ProcImprimirRel FormulaRelatorio, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_identificacao_personalizada_Click()
On Error GoTo tratar_erro

ProcAbreModuloEtiqueta True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEstoque_Click()
On Error GoTo tratar_erro

frmestoque_item_relat.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdEtiqueta_Click()
On Error GoTo tratar_erro

ProcAbreModuloEtiqueta False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbreModuloEtiqueta(Personalizado As Boolean)
On Error GoTo tratar_erro

'With frmestoque_item
    If RE = 0 Then
        USMsgBox ("Informe o RE na lista antes de abrir o menu para emissão de etiquetas."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
'End With

Inspecao_recebimento = False
Estoque_recebimento = False
Faturamento = False
Permitido = Personalizado
frmestoque_item_imprimir_etiqueta.Show 1

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
