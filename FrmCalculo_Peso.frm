VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form FrmCalculo_Peso 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculadora para cálculo de peso"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7740
   Icon            =   "FrmCalculo_Peso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1530
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   1485
      Left            =   55
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   0
      Width           =   7635
      Begin VB.TextBox txtcodigo 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   2115
      End
      Begin VB.TextBox txtun 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   390
         Width           =   735
      End
      Begin VB.ComboBox cmbunkg 
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
         ItemData        =   "FrmCalculo_Peso.frx":030A
         Left            =   4890
         List            =   "FrmCalculo_Peso.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Unidade por kilograma."
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtquantidade 
         Alignment       =   1  'Alinhar à Direita
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
         Height          =   315
         Left            =   2670
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "0,00000"
         ToolTipText     =   "Quantidade."
         Top             =   1020
         Width           =   1155
      End
      Begin VB.TextBox txtpeso 
         Alignment       =   1  'Alinhar à Direita
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
         Height          =   315
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "0,00000"
         ToolTipText     =   "Kilograma por unidade."
         Top             =   390
         Width           =   1410
      End
      Begin VB.TextBox txtdimensao 
         Alignment       =   1  'Alinhar à Direita
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
         Height          =   315
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "0,00000"
         ToolTipText     =   "Dimensão a ser utilizada por peça."
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txtkgpc 
         Alignment       =   1  'Alinhar à Direita
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
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso por peça."
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox txtpesototal 
         Alignment       =   1  'Alinhar à Direita
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
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Peso total."
         Top             =   1020
         Width           =   1335
      End
      Begin DrawSuite2022.USButton Cmd_carregar 
         Height          =   315
         Left            =   5370
         TabIndex        =   8
         ToolTipText     =   "Carregar peso total na quantidade."
         Top             =   1020
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Carregar (F3)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         HandPointer     =   0   'False
         PicAlign        =   8
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Código interno"
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
         Height          =   195
         Left            =   712
         TabIndex        =   18
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Alinhar à Direita
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Quant."
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
         Left            =   2970
         TabIndex        =   17
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Un."
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
         Height          =   195
         Left            =   2550
         TabIndex        =   16
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Kg/unidade"
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
         Left            =   3285
         TabIndex        =   15
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   4650
         TabIndex        =   14
         Top             =   480
         Width           =   105
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Centralizar
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Dim. / mm"
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
         Left            =   6000
         TabIndex        =   13
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Kg/pç"
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
         Left            =   1620
         TabIndex        =   12
         Top             =   810
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Peso total"
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
         Left            =   4200
         TabIndex        =   11
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Un/Kg"
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
         Left            =   5175
         TabIndex        =   10
         Top             =   180
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmCalculo_Peso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbunkg_Click()
On Error GoTo tratar_erro

If cmbunkg = "Mt²" Then
    If txtun = "MT" Then Label24.Caption = "Area / mt" Else Label24.Caption = "Area / mm"
Else
    Label24.Caption = "Dim. / mm"
End If
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_carregar_Click()
On Error GoTo tratar_erro

If Estoque_recebimento = True Then
    frmEstoque_Recebimento.txtQuantidade = txtpesototal
ElseIf Compras_Requisicao = True Then
        frmCompras_Requisicao.txtQS_est = txtpesototal
    ElseIf Compras_Cotacao = True Then
            frmCompras_reqcot_abrir.txtQtde = txtpesototal
        ElseIf Compras_Pedido = True Then
                frmCompras_Pedido.txtQuantidade = txtpesototal
            ElseIf Vendas_Proposta = True Then
                    frmVendas_proposta.txtQuantidade = txtpesototal
                Else
                    frmVendas_PI.txtQuantidade = txtpesototal
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

txtCodigo = TBProduto!Desenho
txtun.Text = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
If IsNull(TBProduto!Un_Kg) = False Then cmbunkg.Text = TBProduto!Un_Kg
txtpeso.Text = IIf(IsNull(TBProduto!peso_metro), "", Format(TBProduto!peso_metro, "###,##0.0000000000"))
txtQuantidade.Text = "1,00000"
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: Cmd_carregar_Click
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdimensao_Change()
On Error GoTo tratar_erro

If txtdimensao.Text <> "" Then
    VerifNumero = txtdimensao.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtdimensao.Text = ""
        txtdimensao.SetFocus
        Exit Sub
    End If
End If
ProcCalculaPeso
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaPeso()
On Error GoTo tratar_erro

If txtpeso.Text <> "" And cmbunkg.Text <> "" And txtdimensao.Text <> "" And txtQuantidade.Text <> "" Then
    If cmbunkg.Text = "Mt/L" Then txtkgpc.Text = Format(txtpeso.Text / 1000 * txtdimensao, "###,##0.0000000000")
    If cmbunkg.Text = "Pç" Then txtkgpc.Text = Format(txtpeso.Text, "###,##0.0000000000")
    If cmbunkg.Text = "Mt²" Then txtkgpc.Text = Format(((txtdimensao * txtpeso) / 1000) / 1000, "###,##0.0000000000")
    If cmbunkg.Text = "N/a" Then txtkgpc.Text = Format(0, "###,##0.0000000000")
    If txtdimensao.Text = "" Then txtdimensao.Text = Format(0, "###,##0.0000000000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcalculaPesoTotal()
On Error GoTo tratar_erro

If txtkgpc.Text <> "" And txtQuantidade <> "" Then
    txtpesototal = Format(txtkgpc.Text * txtQuantidade.Text, "###,##0.0000000000")
Else
    txtpesototal = "0,0000"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaValor()
On Error GoTo tratar_erro

ProcCalculaPeso
ProcalculaPesoTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdimensao_LostFocus()
On Error GoTo tratar_erro

txtdimensao.Text = Format(txtdimensao.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtkgpc_Change()
On Error GoTo tratar_erro

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtkgpc_LostFocus()
On Error GoTo tratar_erro

If txtkgpc.Text <> "" Then
    VerifNumero = txtkgpc.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtkgpc.Text = ""
        txtkgpc.SetFocus
        Exit Sub
    End If
    txtkgpc.Text = Format(txtkgpc.Text, "###,##0.0000000000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpeso_Change()
On Error GoTo tratar_erro

If txtpeso.Text <> "" Then
    VerifNumero = txtpeso.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtpeso.Text = ""
        txtpeso.SetFocus
        Exit Sub
    End If
End If
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpeso_LostFocus()
On Error GoTo tratar_erro

txtpeso.Text = Format(txtpeso.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpesototal_Change()
On Error GoTo tratar_erro

ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantidade_Change()
On Error GoTo tratar_erro

If txtQuantidade.Text <> "" Then
    VerifNumero = txtQuantidade.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantidade.Text = ""
        txtQuantidade.SetFocus
        Exit Sub
    End If
End If
ProcVerificaValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantidade_LostFocus()
On Error GoTo tratar_erro

txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.0000000000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
