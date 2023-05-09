VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_EncerrarOrdem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "       PCP - Encerrar ordem de produção"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProd_EncerrarOrdem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   13
      Top             =   5130
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1665
      Left            =   420
      TabIndex        =   6
      Top             =   510
      Width           =   5175
      Begin DrawSuite2022.USAlphaImage USAlphaImage1 
         Height          =   540
         Left            =   2250
         Top             =   300
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "frmProd_EncerrarOrdem.frx":000C
         Props           =   5
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmProd_EncerrarOrdem.frx":0B03
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   240
         TabIndex        =   7
         Top             =   870
         Width           =   4665
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções de estoque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   420
      TabIndex        =   8
      Top             =   3120
      Width           =   2925
      Begin VB.CheckBox chkprod 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrar com item no estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   510
         Value           =   1  'Marcado
         Width           =   2625
      End
      Begin VB.CheckBox chkReq 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Baixar requisições do estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Value           =   1  'Marcado
         Width           =   2475
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data encerramento ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3360
      TabIndex        =   10
      Top             =   3120
      Width           =   2235
      Begin MSComCtl2.DTPicker dtEncerramento 
         Height          =   315
         Left            =   300
         TabIndex        =   4
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   321191937
         CurrentDate     =   43594
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   420
      TabIndex        =   11
      Top             =   2130
      Width           =   5175
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   240
         Left            =   585
         Top             =   210
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   423
         Caption         =   "Aprovado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Aprovado"
      End
      Begin DrawSuite2022.USTextBoxEx txtQuantProd 
         Height          =   375
         Left            =   300
         TabIndex        =   0
         Top             =   420
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Alignment       =   2
         AllowNegativeNumber=   0   'False
         AutoFormatDate  =   -1  'True
         CurrencyChar    =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
         FormatType      =   2
         MaskType        =   1
         MaxLength       =   29
         NumberOnly      =   -1  'True
         Text            =   "0,00"
      End
      Begin DrawSuite2022.USTextBoxEx txtQuantNC 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   420
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Alignment       =   2
         AutoFormatDate  =   -1  'True
         CurrencyChar    =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ForeColorOver   =   128
         FormatType      =   2
         MaskType        =   1
         MaxLength       =   29
         NumberOnly      =   -1  'True
         Text            =   "0,00"
      End
      Begin DrawSuite2022.USTextBoxEx txtQuantTotal 
         Height          =   375
         Left            =   3540
         TabIndex        =   12
         Top             =   420
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Alignment       =   2
         AutoFormatDate  =   -1  'True
         CurrencyChar    =   ""
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatType      =   2
         Locked          =   -1  'True
         MaskType        =   1
         MaxLength       =   29
         NumberOnly      =   -1  'True
         Text            =   "0,00"
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   240
         Left            =   1980
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Caption         =   "Não conforme"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
         NoHTMLCaption   =   "Não conforme"
      End
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   240
         Left            =   3780
         Top             =   210
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   423
         Caption         =   "Produzido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Produzido"
      End
   End
   Begin DrawSuite2022.USButton cmdEncerrarordem 
      Height          =   825
      Left            =   390
      TabIndex        =   5
      Top             =   4140
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1455
      DibPicture      =   "frmProd_EncerrarOrdem.frx":0BB2
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Encerrar ordem de produção"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      PicSize         =   5
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   767
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
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmProd_EncerrarOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEncerrarordem_Click()
On Error GoTo tratar_erro
Dim QuantProd As Double
Dim QuantNC As Double


'==================================
' NOVA FUNÇÃO BAIXAR ESTOQUE 2019
'==================================
With frmprod
 If .cmbstatus.Text = "Aberta" Then
 
    
    DataOrdem = dtEncerramento.Value
    
      If Not IsDate(DataOrdem) Then
      USMsgBox "Data informada para encerramento da ordem é inválida, por favor informar a data correta", vbInformation, "CAPRIND v5.0"
      Exit Sub
      End If
    
 End If
'================================
' NOVA FUNÇÃO BAIXAR ESTOQUE
'==================================
'   Concluir ordem de produção e colocar data de conclusão
    If .cmbstatus.Text <> "Concluída" Then
      If chkReq.Value = 1 Then
       .ProcQTBaixar_estoque_ordem
       USMsgBox "Requisições baixadas do estoque com sucesso!", vbInformation, "CAPRIND v5.0"
      End If
      If chkprod.Value = 1 Then
       .ProcEntradaEstoqueOrdem
       USMsgBox "Entrada do item no estoque com sucesso!", vbInformation, "CAPRIND v5.0"
      End If
    End If
    
QuantProd = txtQuantProd.Text
QuantNC = txtQuantNC.Text

'QuantProd = 1
'QuantNC = 1

'Encerra ordem de produção
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from Producao where Ordem = " & NOrdem, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
TBEstoque!pronta = "SIM"
TBEstoque!DataEntrega = DataOrdem
TBEstoque!status = "Concluida"
TBEstoque!QuantProd = txtQuantProd.Text
TBEstoque!QuantNC = txtQuantNC
TBEstoque.Update
End If

'Encerra as ordens de serviço
Conexao.Execute "update Ordemservico set pronto='SIM' , Dataconclusao = '" & DataOrdem & "' , Status = 'Concluída' where Ordem = " & NOrdem

'Atualiza lista de ordens de produção
.atualiza_lista_ordens (1)
End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

dtEncerramento.Value = frmprod.mskprazofina.Value

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantNC_LostFocus()
On Error GoTo tratar_erro
Dim Prod As Double
Dim QTNC As Double

Prod = txtQuantProd.Text
QTNC = txtQuantNC.Text

If IsNumeric(Prod) And IsNumeric(QTNC) Then
txtQuantTotal.Text = Prod + QTNC
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantProd_LostFocus()
On Error GoTo tratar_erro
Dim Prod As Double
Dim QTNC As Double

Prod = txtQuantProd.Text
QTNC = txtQuantNC.Text

If IsNumeric(Prod) And IsNumeric(QTNC) Then
txtQuantTotal.Text = Prod + QTNC
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
